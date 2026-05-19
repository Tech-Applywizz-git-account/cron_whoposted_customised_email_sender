const cron = require("node-cron");
const { createClient } = require("@supabase/supabase-js");
const { sendActiveUpdateEmail, sendRenewalEmail, sendUpsellEmail, sendCampaignEmail } = require("./mailer");
const { DateTime } = require("luxon");
require("dotenv").config();

// ============================================================
//  TABLE CONFIGURATION
// ============================================================
// Dummy tables for testing:
const USERS_TABLE = "whopost_users_dummy";
const TRANSACTIONS_TABLE = "whoposted_transactions_dummy";

// Production tables (swap when ready):
// const USERS_TABLE = "whopost_users";
// const TRANSACTIONS_TABLE = "whoposted_transactions";

const JOBS_TABLE = "daily_linkedin_jobs_report";

// ============================================================
//  ENVIRONMENT VALIDATION
// ============================================================
const requiredEnv = [
    "SUPABASE_URL",
    "SUPABASE_SERVICE_ROLE_KEY",
    "MS_CLIENT_ID",
    "MS_TENANT_ID",
    "MS_CLIENT_SECRET",
    "MS_SENDER_EMAIL"
];
const missingEnv = requiredEnv.filter(key => !process.env[key]);
if (missingEnv.length > 0) {
    console.error(`CRITICAL ERROR: Missing environment variables: ${missingEnv.join(", ")}`);
    process.exit(1);
}

// ============================================================
//  SUPABASE CLIENT
// ============================================================
const supabase = createClient(
    process.env.SUPABASE_URL,
    process.env.SUPABASE_SERVICE_ROLE_KEY,
    { auth: { persistSession: false }, realtime: { enabled: false } }
);

// ============================================================
//  EXECUTION LOCK — prevents overlapping 5-min cron runs
// ============================================================
let isRunning = false;

// ============================================================
//  HELPERS
// ============================================================
const sleep = (ms) => new Promise(resolve => setTimeout(resolve, ms));

async function downloadAttachment(url, retries = 3, delay = 2000) {
    for (let i = 1; i <= retries; i++) {
        try {
            const response = await fetch(url);
            if (response.status === 502 || !response.ok) {
                throw new Error(`HTTP ${response.status}: ${response.statusText}`);
            }
            const buffer = Buffer.from(await response.arrayBuffer());
            return buffer.toString("base64");
        } catch (err) {
            console.warn(`[Attachment] Attempt ${i}/${retries} failed to download ${url}: ${err.message}`);
            if (i === retries) {
                console.error(`[Attachment] Max retries reached. Failed to download ${url}.`);
                return null;
            }
            await sleep(delay);
        }
    }
    return null;
}

// ============================================================
//  CORE EMAIL DISPATCH — sends one campaign to all matching users
// ============================================================
async function dispatchCampaign(campaign, users, userExpiries, jobsToInclude) {
    const segment = campaign.segment_type;
    const now = new Date();
    let successCount = 0;
    let failCount = 0;

    // Pre-download attachments once per campaign
    const preparedAttachments = [];
    if (Array.isArray(campaign.attachments)) {
        for (const att of campaign.attachments) {
            console.log(`[Attachment] Downloading: ${att.name}`);
            const base64 = await downloadAttachment(att.url);
            if (base64) {
                preparedAttachments.push({
                    name: att.name,
                    url: att.url,
                    contentType: att.type,
                    contentBytes: base64
                });
            }
        }
    }

    for (const user of users) {
        const userEmail = user.email;
        const clientName = user.full_name || "there";
        const latestExpiry = userExpiries.get(userEmail);

        // Determine user's segment type
        let userType = "free";
        if (latestExpiry && latestExpiry > now) userType = "active";
        else if (latestExpiry) userType = "expired";

        // Only send to the matching segment
        if (userType !== segment) continue;

        try {
            // Create tracking record
            const { data: trackRecord, error: trackError } = await supabase
                .from("email_tracking")
                .insert({ user_email: userEmail, user_type: userType })
                .select("id")
                .single();

            if (trackError) throw new Error(`Tracking Insert Error: ${trackError.message}`);
            const trackingId = trackRecord.id;

            // Send the custom campaign email
            await sendCampaignEmail(
                userEmail,
                clientName,
                campaign.subject,
                campaign.body,
                trackingId,
                preparedAttachments,
                jobsToInclude,
                userType,
                {
                    enabled: campaign.cta_enabled,
                    text: campaign.cta_text,
                    link: campaign.cta_link
                },
                campaign.jobs_enabled,
                campaign.jobs_title,
                campaign.jobs_footer
            );

            successCount++;
            // Rate limiting: ~30 emails/min
            await sleep(2000);
        } catch (mailErr) {
            console.error(`[Dispatch] Failed for ${userEmail}:`, mailErr.message);
            failCount++;
        }
    }

    return { successCount, failCount };
}

// ============================================================
//  MAIN CRON PROCESSOR — runs every 5 minutes
// ============================================================
async function processDueCampaigns() {
    // Execution lock — bail if previous run is still in progress
    if (isRunning) {
        console.log(`[Lock] Previous run still in progress. Skipping this tick.`);
        return;
    }

    isRunning = true;
    const istNow = DateTime.now().setZone("Asia/Kolkata");
    const logTime = istNow.toFormat("yyyy-MM-dd HH:mm:ss");
    console.log(`\n[${logTime} IST] 5-minute cron triggered.`);

    try {
        // ── Step 1: Fetch due campaigns (IST-aware comparison) ──────────────
        // Finds campaigns where the scheduled date+time (interpreted in IST) is ≤ now
        // Prioritizes: failed campaigns first (within last 30 min, max 3 retries), then ready ones
        const { data: campaigns, error: fetchError } = await supabase.rpc("get_due_campaigns");

        if (fetchError) throw new Error(`Campaign Fetch Error: ${fetchError.message}`);

        if (!campaigns || campaigns.length === 0) {
            console.log(`[${logTime} IST] No due campaigns. Standing by.`);
            return;
        }

        console.log(`[${logTime} IST] Found ${campaigns.length} due campaign(s) to dispatch.`);

        // ── Step 2: Fetch 3 latest jobs from last 24h ────────────────────────
        const last24Hours = istNow.minus({ hours: 24 }).toISO();
        const { data: latestJobs } = await supabase
            .from(JOBS_TABLE)
            .select("job_title, company, poster_profile_url")
            .gte("created_at", last24Hours)
            .order("created_at", { ascending: false })
            .limit(3);
        const jobsToInclude = latestJobs || [];

        // ── Step 3: Fetch all users and subscription data once ──────────────
        const { data: users, error: usersError } = await supabase
            .from(USERS_TABLE)
            .select("email, full_name");
        if (usersError) throw new Error(`Users Fetch Error: ${usersError.message}`);

        const { data: transactions, error: transError } = await supabase
            .from(TRANSACTIONS_TABLE)
            .select("user_email, expiry_date")
            .order("expiry_date", { ascending: false });
        if (transError) throw new Error(`Transactions Fetch Error: ${transError.message}`);

        // Build expiry lookup map (most recent expiry per user)
        const userExpiries = new Map();
        transactions.forEach(t => {
            if (!userExpiries.has(t.user_email)) {
                userExpiries.set(t.user_email, new Date(t.expiry_date));
            }
        });

        // ── Step 4: Process each due campaign ───────────────────────────────
        for (const campaign of campaigns) {
            const campaignId = campaign.id;
            const campaignLabel = `[${campaign.segment_type.toUpperCase()} @ ${campaign.scheduled_date} ${campaign.scheduled_time} IST]`;

            console.log(`\n${campaignLabel} Starting dispatch (retry_count: ${campaign.retry_count})...`);

            // Mark as 'picked' BEFORE sending — acts as a distributed lock
            const { error: pickError } = await supabase
                .from("campaigns")
                .update({
                    status: "picked",
                    retry_count: (campaign.retry_count || 0) + 1,
                    updated_at: new Date().toISOString()
                })
                .eq("id", campaignId);

            if (pickError) {
                console.error(`${campaignLabel} Failed to mark as picked:`, pickError.message);
                continue;
            }

            try {
                const { successCount, failCount } = await dispatchCampaign(
                    campaign,
                    users,
                    userExpiries,
                    jobsToInclude
                );

                console.log(`${campaignLabel} Dispatch complete — Sent: ${successCount}, Failed: ${failCount}`);

                if (successCount === 0) {
                    // Nothing sent at all — safe to retry (no duplicates risk)
                    await supabase
                        .from("campaigns")
                        .update({ status: "failed", updated_at: new Date().toISOString() })
                        .eq("id", campaignId);

                    console.log(`${campaignLabel} ⚠️ 0 emails sent. Status set to 'failed'. Will retry if within 30-min window and retry_count < 3.`);
                } else {
                    // At least some emails sent — mark completed, do NOT retry
                    // (retrying would send duplicates to users who already received the email)
                    await supabase
                        .from("campaigns")
                        .update({ status: "completed", updated_at: new Date().toISOString() })
                        .eq("id", campaignId);

                    if (failCount > 0) {
                        console.log(`${campaignLabel} ✅ Completed with ${failCount} individual failure(s). No retry (would cause duplicates). Check logs above for specific failed addresses.`);
                    } else {
                        console.log(`${campaignLabel} ✅ All ${successCount} emails sent successfully.`);
                    }
                }

            } catch (dispatchErr) {
                // Critical error — entire dispatch threw before sending anything
                console.error(`${campaignLabel} ❌ Critical dispatch error:`, dispatchErr.message);

                await supabase
                    .from("campaigns")
                    .update({ status: "failed", updated_at: new Date().toISOString() })
                    .eq("id", campaignId);

                console.log(`${campaignLabel} Status set to 'failed'. Will retry if within 30-min window and retry_count < 3.`);
            }
        }

    } catch (err) {
        console.error(`[FATAL CRON ERROR]:`, err.message);
    } finally {
        // Always release the lock
        isRunning = false;
        console.log(`\n[${istNow.toFormat("HH:mm:ss")} IST] Cron run complete. Lock released.`);
    }
}

// ============================================================
//  CRON SCHEDULE — every 5 minutes, IST timezone
// ============================================================
if (!process.env.VERCEL) {
    cron.schedule("*/5 * * * *", () => {
        processDueCampaigns();
    }, {
        scheduled: true,
        timezone: "Asia/Kolkata"
    });

    console.log("=================================================");
    console.log("  WhoPosted Dynamic Campaign Mailer — LIVE");
    console.log("  Schedule: Every 5 minutes (IST)");
    console.log("  Mode: Dynamic date+time per campaign");
    console.log("  Failover: Auto-retry failed campaigns (max 3x / 30min)");
    console.log("  Lock: Single-instance execution guard active");
    console.log(`  Users Table: ${USERS_TABLE}`);
    console.log("=================================================\n");
} else {
    console.log("[Serverless] Running in serverless context. Internal cron scheduler disabled.");
}

module.exports = { processDueCampaigns };
