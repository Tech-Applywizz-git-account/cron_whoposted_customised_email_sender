const { processDueCampaigns } = require("../index");

module.exports = async (req, res) => {
    // Optional check for Vercel Cron authorization secret
    if (process.env.CRON_SECRET && req.headers.authorization !== `Bearer ${process.env.CRON_SECRET}`) {
        return res.status(401).json({ error: "Unauthorized" });
    }

    try {
        console.log("[Vercel Cron] Starting campaign processing...");
        await processDueCampaigns();
        console.log("[Vercel Cron] Campaign processing finished successfully.");
        return res.status(200).json({ success: true, message: "Campaign cron completed successfully" });
    } catch (err) {
        console.error("[Vercel Cron] Fatal error during campaign run:", err.message);
        return res.status(500).json({ success: false, error: err.message });
    }
};
