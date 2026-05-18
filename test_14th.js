const { createClient } = require("@supabase/supabase-js");
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

async function runTestFor14th() {
    const TEST_DATE = "2026-05-14"; 
    console.log(`\n🚀 STARTING PRODUCTION TEST RUN FOR TODAY: ${TEST_DATE}`);
    console.log(`Using Dummy Tables: ${USERS_TABLE}, ${TRANSACTIONS_TABLE}\n`);

    try {
        // 1. Fetch campaigns for today
        const { data: campaigns, error: campaignError } = await supabase
            .from("campaigns")
            .select("*")
            .eq("scheduled_date", TEST_DATE)
            .eq("status", "ready")
            .eq("is_deleted", false);

        if (campaignError) throw campaignError;
        
        console.log(`✅ Found ${campaigns?.length || 0} custom campaigns for today.`);
        if (campaigns) {
            campaigns.forEach(c => console.log(`   - [${c.segment_type.toUpperCase()}] Campaign exists`));
        }

        // 2. Prepare campaigns
        const campaignMap = new Map();
        const { sendCampaignEmail, sendActiveUpdateEmail, sendRenewalEmail, sendUpsellEmail } = require("./mailer");

        if (campaigns) {
            for (const c of campaigns) {
                console.log(`Preparing ${c.segment_type.toUpperCase()} campaign data...`);
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

        // 4. Fetch dummy users (testing with a small sample)
        const { data: users } = await supabase.from(USERS_TABLE).select("email, full_name").limit(10);
        
        // 5. Fetch dummy transactions for expiries
        const { data: transactions } = await supabase.from(TRANSACTIONS_TABLE).select("user_email, expiry_date");
        const userExpiries = new Map();
        transactions?.forEach(t => userExpiries.set(t.user_email, new Date(t.expiry_date)));

        console.log(`\n📧 Processing ${users.length} sample users...\n`);

        for (const user of users) {
            const latestExpiry = userExpiries.get(user.email);
            const now = new Date();
            let userType = "free";
            if (latestExpiry && latestExpiry > now) userType = "active";
            else if (latestExpiry) userType = "expired";

            const campaignData = campaignMap.get(userType);
            
            if (campaignData) {
                console.log(`✉️  [CUSTOM] Sending to ${userType.toUpperCase()} user: ${user.email}`);
                await sendCampaignEmail(
                    user.email,
                    user.full_name || "there",
                    campaignData.subject,
                    campaignData.body,
                    null, 
                    campaignData.preparedAttachments,
                    jobs,
                    userType,
                    {
                        enabled: campaignData.cta_enabled,
                        text: campaignData.cta_text,
                        link: campaignData.cta_link
                    },
                    campaignData.jobs_enabled,
                    campaignData.jobs_title,
                    campaignData.jobs_footer
                );
            } else {
                console.log(`✉️  [DEFAULT] Sending to ${userType.toUpperCase()} user: ${user.email}`);
                if (userType === "active") await sendActiveUpdateEmail(user.email, user.full_name || "there", jobs);
                else if (userType === "expired") await sendRenewalEmail(user.email, user.full_name || "there", jobs);
                else await sendUpsellEmail(user.email, user.full_name || "there", jobs);
            }
            console.log(`   ✅ Success`);
        }

        console.log(`\n✨ TEST COMPLETED! Check your inbox (if you used your own emails in dummy table).`);

    } catch (error) {
        console.error("\n💥 FATAL ERROR:", error.message);
    }
}

runTestFor14th();
