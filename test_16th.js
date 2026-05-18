const { createClient } = require("@supabase/supabase-js");
const { DateTime } = require("luxon");
require("dotenv").config();

// Dummy Table Names for Testing
const USERS_TABLE = "whopost_users_dummy";
const TRANSACTIONS_TABLE = "whoposted_transactions_dummy";
const JOBS_TABLE = "daily_linkedin_jobs_report";

// Initialize Supabase
const supabase = createClient(process.env.SUPABASE_URL, process.env.SUPABASE_SERVICE_ROLE_KEY);

async function downloadAttachment(url) {
    try {
        const response = await fetch(url);
        if (!response.ok) throw new Error(`Failed to download: ${response.statusText}`);
        const arrayBuffer = await response.arrayBuffer();
        return Buffer.from(arrayBuffer).toString('base64');
    } catch (err) {
        console.error(`Error downloading attachment:`, err.message);
        return null;
    }
}

async function runTestFor16th() {
    const TEST_DATE = "2026-05-19"; 
    console.log(`\n🚀 STARTING TEST RUN FOR DATE: ${TEST_DATE}`);
    console.log(`Using Dummy Tables: ${USERS_TABLE}, ${TRANSACTIONS_TABLE}\n`);

    try {
        // 1. Fetch campaigns for the 16th
        const { data: campaigns, error: campaignError } = await supabase
            .from("campaigns")
            .select("*")
            .eq("scheduled_date", TEST_DATE)
            .eq("status", "ready")
            .eq("is_deleted", false);

        if (campaignError) throw campaignError;
        
        if (!campaigns || campaigns.length === 0) {
            console.log("ℹ️ No custom campaigns found for the 16th. Will use DEFAULT templates.");
        } else {
            console.log(`✅ Found ${campaigns.length} custom campaigns for the 16th.`);
        }

        // 2. Prepare campaigns
        const campaignMap = new Map();
        const { sendCampaignEmail, sendActiveUpdateEmail, sendRenewalEmail, sendUpsellEmail } = require("./mailer");

        if (campaigns) {
            for (const c of campaigns) {
                console.log(`Preparing ${c.segment_type.toUpperCase()} campaign...`);
                const preparedAttachments = [];
                if (c.attachments && Array.isArray(c.attachments)) {
                    for (const att of c.attachments) {
                        const base64 = await downloadAttachment(att.url);
                        if (base64) {
                            preparedAttachments.push({
                                name: att.name,
                                url: att.url,
                                contentType: att.type || "application/pdf",
                                contentBytes: base64
                            });
                        }
                    }
                }
                campaignMap.set(c.segment_type, { ...c, preparedAttachments });
            }
        }

        // 3. Fetch latest jobs
        const { data: latestJobs } = await supabase.from(JOBS_TABLE).select("*").limit(3).order('created_at', { ascending: false });
        const jobs = latestJobs || [];

        // 4. Fetch dummy users
        const { data: users } = await supabase.from(USERS_TABLE).select("email, full_name");
        
        // 5. Fetch dummy transactions for expiries
        const { data: transactions } = await supabase.from(TRANSACTIONS_TABLE).select("user_email, expiry_date");
        const userExpiries = new Map();
        transactions?.forEach(t => userExpiries.set(t.user_email, new Date(t.expiry_date)));

        console.log(`\n📧 Processing ${users.length} users...\n`);

        for (const user of users) {
            const latestExpiry = userExpiries.get(user.email);
            const now = new Date();
            let userType = "free";
            if (latestExpiry && latestExpiry > now) userType = "active";
            else if (latestExpiry) userType = "expired";

            const campaign = campaignMap.get(userType);
            
            if (campaign) {
                console.log(`✉️  Sending CUSTOM ${userType.toUpperCase()} email to: ${user.email}`);
                await sendCampaignEmail(
                    user.email,
                    user.full_name || "there",
                    campaign.subject,
                    campaign.body,
                    null, 
                    campaign.preparedAttachments,
                    jobs,
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
            } else {
                console.log(`✉️  Sending DEFAULT ${userType.toUpperCase()} email to: ${user.email}`);
                if (userType === "active") await sendActiveUpdateEmail(user.email, user.full_name || "there", jobs);
                else if (userType === "expired") await sendRenewalEmail(user.email, user.full_name || "there", jobs);
                else await sendUpsellEmail(user.email, user.full_name || "there", jobs);
            }
            console.log(`✅ Success`);
        }

        console.log(`\n✨ TEST COMPLETED!`);

    } catch (error) {
        console.error("\n💥 FATAL ERROR:", error.message);
    }
}

runTestFor16th();
