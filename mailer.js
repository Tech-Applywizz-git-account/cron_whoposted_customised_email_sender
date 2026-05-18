require("dotenv").config();

async function getMicrosoftAccessToken() {
    const clientId = process.env.MS_CLIENT_ID;
    const tenantId = process.env.MS_TENANT_ID;
    const clientSecret = process.env.MS_CLIENT_SECRET;

    if (!clientId || !tenantId || !clientSecret) {
        throw new Error("Microsoft 365 credentials missing in environment.");
    }

    const url = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`;
    const params = new URLSearchParams({
        client_id: clientId,
        client_secret: clientSecret,
        scope: "https://graph.microsoft.com/.default",
        grant_type: "client_credentials",
    });

    const response = await fetch(url, {
        method: "POST",
        headers: { "Content-Type": "application/x-www-form-urlencoded" },
        body: params.toString(),
    });

    const data = await response.json();
    if (data.error) {
        throw new Error(`MS Graph Token Error: ${data.error_description || data.error}`);
    }

    return data.access_token;
}

// Reusable Job Card Component for Emails
function generateJobsHtml(jobs, isBlurry = false) {
    if (!jobs || jobs.length === 0) return '';

    return `
        <div style="margin: 24px 0; border-radius: 16px; border: 1px solid #e2e8f0; overflow: hidden; background: #ffffff;">
            <div style="padding: 8px 16px;">
                ${jobs.map(job => `
                    <div style="padding: 16px 0; border-bottom: ${jobs.indexOf(job) === jobs.length - 1 ? 'none' : '1px solid #f1f5f9'};">
                        <div style="font-weight: 700; color: #0f172a; font-size: 15px;">${job.job_title}</div>
                        <div style="color: #64748b; font-size: 13px; margin-bottom: 8px;">at ${job.company}</div>
                        ${isBlurry 
                            ? `<div style="background: #f1f5f9; height: 12px; width: 140px; border-radius: 4px; filter: blur(4px);"></div>`
                            : `<a href="${job.poster_profile_url}" style="color: #533afd; text-decoration: none; font-size: 13px; font-weight: 600;">View Recruiter's Profile →</a>`
                        }
                    </div>
                `).join('')}
            </div>
        </div>
    `;
}

// Helper to add the invisible tracking pixel
function getTrackingPixelHtml(trackingId) {
    if (!trackingId) return '';
    // Replace this with your actual production domain
    const baseUrl = process.env.TRACKING_BASE_URL || "https://www.whopostedai.com/api/track-email";
    return `<img src="${baseUrl}?id=${trackingId}" width="1" height="1" style="display:none !important;" />`;
}

async function sendMail(to, subject, htmlContent, attachments = []) {
    const accessToken = await getMicrosoftAccessToken();
    const senderEmail = process.env.MS_SENDER_EMAIL || "krish@whopostedai.com";

    const message = {
        subject: subject,
        body: { contentType: "Html", content: htmlContent },
        toRecipients: [{ emailAddress: { address: to } }],
    };

    if (attachments && attachments.length > 0) {
        message.attachments = attachments.map(att => ({
            "@odata.type": "#microsoft.graph.fileAttachment",
            name: att.name,
            contentType: att.contentType || "application/octet-stream",
            contentBytes: att.contentBytes // This must be base64 string
        }));
    }

    const response = await fetch(`https://graph.microsoft.com/v1.0/users/${senderEmail}/sendMail`, {
        method: "POST",
        headers: { 
            Authorization: `Bearer ${accessToken}`, 
            "Content-Type": "application/json" 
        },
        body: JSON.stringify({ message }),
    });

    if (!response.ok) {
        const error = await response.json();
        throw new Error(`Graph API error for ${to}: ${JSON.stringify(error)}`);
    }
    console.log(`Email sent successfully to ${to}`);
}

const baseStyle = `font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; max-width: 600px; margin: 0 auto; padding: 40px; border: 1px solid #f1f5f9; border-radius: 24px; color: #1e293b;`;
const headerHtml = `
    <div style="text-align: center; margin-bottom: 32px;">
        <h1 style="margin: 0; color: #0f172a; font-size: 28px; font-weight: 900; letter-spacing: -1px;">Who<span style="color: #533afd;">Posted</span></h1>
    </div>
`;
const footerHtml = `
    <div style="margin-top: 40px; padding-top: 24px; border-top: 1px solid #f1f5f9;">
        <p style="margin: 0; color: #0f172a; font-weight: 700;">– Team WhoPosted</p>
    </div>
    <div style="margin-top: 40px; padding-top: 24px; border-top: 1px solid #f1f5f9; text-align: center;">
        <p style="color: #cbd5e1; font-size: 11px; margin: 8px 0 0 0;">&copy; ${new Date().getFullYear()} WhoPosted AI. All rights reserved.</p>
    </div>
`;

// 1. ACTIVE USERS EMAIL
async function sendActiveUpdateEmail(to, clientName, jobs = [], trackingId = null) {
    const htmlContent = `
        <div style="${baseStyle}">
            ${headerHtml}
            <h2 style="font-weight: 900; font-size: 24px; color: #533afd; margin-bottom: 16px;">New Hiring Connections Added</h2>
            <p style="font-size: 16px; line-height: 1.6;">Hi ${clientName},</p>
            <p style="font-size: 16px; line-height: 1.6;">New job postings and hiring connections have been added to your WhoPosted dashboard.</p>
            <p style="font-size: 16px; line-height: 1.6;">These are professionals actively hiring — giving you a chance to connect early and stand out before others apply.</p>
            <p style="font-size: 16px; line-height: 1.6;">We've found 3 recruiters that look like a great fit for you today:</p>            
            ${generateJobsHtml(jobs)}
            <p style="font-size: 16px; line-height: 1.6; font-weight: 600;">Take a few minutes today to reach out and build meaningful connections.</p>
            <div style="background: #533afd; padding: 24px; border-radius: 16px; text-align: center; margin: 32px 0;">
                <a href="https://www.whopostedai.com" style="display: inline-block; padding: 16px 32px; background: #ffffff; color: #533afd !important; text-decoration: none; border-radius: 12px; font-weight: 800; font-size: 14px; text-transform: uppercase;">Go to Dashboard</a>
            </div>
            ${footerHtml}
            ${getTrackingPixelHtml(trackingId)}
        </div>
    `;
    await sendMail(to, "You've got new job posted members", htmlContent);
}

// 2. EXPIRED USERS EMAIL
async function sendRenewalEmail(to, clientName, jobs = [], trackingId = null) {
    const htmlContent = `
        <div style="${baseStyle}">
            ${headerHtml}
            <h2 style="font-weight: 900; font-size: 24px; color: #fda23aff; margin-bottom: 16px;">Restart Your Access to Hiring Contacts</h2>
            <p style="font-size: 16px; line-height: 1.6;">Hi ${clientName},</p>
            <p style="font-size: 16px; line-height: 1.6;">Your WhoPosted subscription has expired.</p>
            <p style="font-size: 16px; line-height: 1.6;">You’re currently missing access to direct hiring contacts and recent job postings where early outreach can make a difference.</p>
            <p style="font-size: 16px; line-height: 1.6;">Here are 3 recruiters active in the last 24 hours:</p>
            ${generateJobsHtml(jobs, true)}
            <p style="font-size: 16px; line-height: 1.6; font-weight: 600;">Renew your access to reconnect with hiring professionals and stay ahead in your job search.</p>
            <div style="background: #0f172a; padding: 24px; border-radius: 16px; text-align: center; margin: 32px 0;">
                <a href="https://www.whopostedai.com/billing" style="display: inline-block; padding: 16px 32px; background: #533afd; color: #ffffff !important; text-decoration: none; border-radius: 12px; font-weight: 800; font-size: 14px; text-transform: uppercase;">Renew Access Now</a>
            </div>
            ${footerHtml}
            ${getTrackingPixelHtml(trackingId)}
        </div>
    `;
    await sendMail(to, "Don't miss out! Your Whoposted subscription has ended", htmlContent);
}

// 3. FREE USERS EMAIL
async function sendUpsellEmail(to, clientName, jobs = [], trackingId = null) {
    const htmlContent = `
        <div style="${baseStyle}">
            ${headerHtml}
            <h2 style="font-weight: 900; font-size: 24px; color: #0f172a; margin-bottom: 16px;">Connect Directly with Hiring Managers</h2>
            <p style="font-size: 16px; line-height: 1.6;">Hi ${clientName},</p>
            <p style="font-size: 16px; line-height: 1.6;">Most candidates rely only on applications — but direct connections often lead to faster responses.</p>
            <p style="font-size: 16px; line-height: 1.6;">WhoPosted gives you access to real job postings along with the LinkedIn profiles of hiring professionals.</p>
            <p style="font-size: 16px; line-height: 1.6;">Here are 3 recruiters active in the last 24 hours:</p>
            ${generateJobsHtml(jobs, true)}
            <p style="font-size: 16px; line-height: 1.6; font-weight: 600;">Upgrade to unlock full access and start connecting directly.</p>
            <div style="background: #533afd; padding: 24px; border-radius: 16px; text-align: center; margin: 32px 0;">
                <a href="https://www.whopostedai.com/pricing" style="display: inline-block; padding: 16px 32px; background: #ffffff; color: #533afd !important; text-decoration: none; border-radius: 12px; font-weight: 800; font-size: 14px; text-transform: uppercase;">Upgrade to Unlock</a>
            </div>
            ${footerHtml}
            ${getTrackingPixelHtml(trackingId)}
        </div>
    `;
    await sendMail(to, "Connect Directly with Hiring Managers", htmlContent);
}

// 4. CUSTOM CAMPAIGN EMAIL (From CMS)
async function sendCampaignEmail(to, clientName, subject, body, trackingId = null, attachments = [], jobs = [], userType = "free", ctaOptions = {}, showJobs = true, jobsTitle = "", jobsFooter = "") {
    const { enabled = false, text = "", link = "" } = ctaOptions;
    
    // Replace {{name}} placeholder if exists
    const personalizedBody = body.replace(/{{name}}/g, clientName);

    // Generate Attachments Links (to show in the body)
    let attachmentsHtml = "";
    if (attachments && attachments.length > 0) {
        attachmentsHtml = `
            <div style="margin-top: 32px; padding: 16px; background: #f8fafc; border: 1px dashed #cbd5e1; border-radius: 16px;">
                <p style="margin: 0 0 12px 0; font-size: 11px; font-weight: 900; color: #64748b; text-transform: uppercase; letter-spacing: 1px;">📎 Attached Documents</p>
                ${attachments.map(att => `
                    <div style="margin-bottom: 8px;">
                        <a href="${att.url}" style="color: #533afd; text-decoration: none; font-size: 14px; font-weight: 700;">${att.name} →</a>
                    </div>
                `).join('')}
            </div>
        `;
    }

    // Generate Custom CTA Button if enabled
    let ctaHtml = "";
    if (enabled) {
        ctaHtml = `
            <div style="background: #533afd; padding: 24px; border-radius: 16px; text-align: center; margin: 32px 0;">
                <a href="${link || '#'}" style="display: inline-block; padding: 16px 32px; background: #ffffff; color: #533afd !important; text-decoration: none; border-radius: 12px; font-weight: 800; font-size: 14px; text-transform: uppercase;">${text || 'Click Here'}</a>
            </div>
        `;
    }

    // Generate Jobs section if enabled
    let jobsHtml = "";
    if (showJobs) {
        const defaultFooter = userType === 'active' 
            ? "Take a few minutes today to reach out and build meaningful connections." 
            : "Renew your access to reconnect with hiring professionals and stay ahead.";

        jobsHtml = `
            <p style="font-size: 11px; font-weight: 900; color: #94a3b8; text-transform: uppercase; letter-spacing: 1px; margin-top: 32px; margin-bottom: 8px;">
                ${jobsTitle || "Recent Job Postings (24h)"}
            </p>
            ${generateJobsHtml(jobs, userType !== "active")}

            ${jobsFooter ? `
            <p style="font-size: 13px; color: #475569; font-weight: 600; margin-top: 24px; white-space: pre-wrap;">
                ${jobsFooter}
            </p>` : ''}
        `;
    }

    const htmlContent = `
        <div style="${baseStyle}">
            ${headerHtml}
            ${personalizedBody}
            
            ${attachmentsHtml}

            ${jobsHtml}

            ${ctaHtml}

            ${footerHtml}
            ${getTrackingPixelHtml(trackingId)}
        </div>
    `;
    await sendMail(to, subject, htmlContent, attachments);
}

module.exports = { 
    sendActiveUpdateEmail, 
    sendRenewalEmail, 
    sendUpsellEmail,
    sendCampaignEmail
};
 




// require("dotenv").config();

// async function getMicrosoftAccessToken() {
//     const clientId = process.env.MS_CLIENT_ID;
//     const tenantId = process.env.MS_TENANT_ID;
//     const clientSecret = process.env.MS_CLIENT_SECRET;

//     if (!clientId || !tenantId || !clientSecret) {
//         throw new Error("Microsoft 365 credentials missing in environment.");
//     }

//     const url = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`;
//     const params = new URLSearchParams({
//         client_id: clientId,
//         client_secret: clientSecret,
//         scope: "https://graph.microsoft.com/.default",
//         grant_type: "client_credentials",
//     });

//     const response = await fetch(url, {
//         method: "POST",
//         headers: { "Content-Type": "application/x-www-form-urlencoded" },
//         body: params.toString(),
//     });

//     const data = await response.json();
//     if (data.error) {
//         throw new Error(`MS Graph Token Error: ${data.error_description || data.error}`);
//     }

//     return data.access_token;
// }

// // Reusable Job Card Component for Emails
// function generateJobsHtml(jobs, isBlurry = false) {
//     if (!jobs || jobs.length === 0) return '';
    
//     return `
//         <div style="margin: 24px 0; border-radius: 16px; border: 1px solid #e2e8f0; overflow: hidden; background: #ffffff;">
//             <div style="padding: 8px 16px;">
//                 ${jobs.map(job => `
//                     <div style="padding: 16px 0; border-bottom: ${jobs.indexOf(job) === jobs.length - 1 ? 'none' : '1px solid #f1f5f9'};">
//                         <div style="font-weight: 700; color: #0f172a; font-size: 15px;">${job.job_title}</div>
//                         <div style="color: #64748b; font-size: 13px; margin-bottom: 8px;">at ${job.company}</div>
//                         ${isBlurry 
//                             ? `<div style="background: #f1f5f9; height: 12px; width: 140px; border-radius: 4px; filter: blur(4px);"></div>`
//                             : `<a href="${job.poster_profile_url}" style="color: #533afd; text-decoration: none; font-size: 13px; font-weight: 600;">View Recruiter's Profile →</a>`
//                         }
//                     </div>
//                 `).join('')}
//             </div>
//         </div>
//     `;
// }

// // Helper to add the invisible tracking pixel
// function getTrackingPixelHtml(trackingId) {
//     if (!trackingId) return '';
//     // Replace this with your actual production domain
//     const baseUrl = process.env.TRACKING_BASE_URL || "https://www.whopostedai.com/api/track-email";
//     return `<img src="${baseUrl}?id=${trackingId}" width="1" height="1" style="display:none !important;" />`;
// }

// async function sendMail(to, subject, htmlContent) {
//     const accessToken = await getMicrosoftAccessToken();
//     const senderEmail = process.env.MS_SENDER_EMAIL || "krish@whopostedai.com";

//     const response = await fetch(`https://graph.microsoft.com/v1.0/users/${senderEmail}/sendMail`, {
//         method: "POST",
//         headers: { 
//             Authorization: `Bearer ${accessToken}`, 
//             "Content-Type": "application/json" 
//         },
//         body: JSON.stringify({
//             message: {
//                 subject: subject,
//                 body: { contentType: "Html", content: htmlContent },
//                 toRecipients: [{ emailAddress: { address: to } }],
//             }
//         }),
//     });

//     if (!response.ok) {
//         const error = await response.json();
//         throw new Error(`Graph API error for ${to}: ${JSON.stringify(error)}`);
//     }
//     console.log(`Email sent successfully to ${to}`);
// }

// const baseStyle = `font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; max-width: 600px; margin: 0 auto; padding: 40px; border: 1px solid #f1f5f9; border-radius: 24px; color: #1e293b;`;
// const headerHtml = `
//     <div style="text-align: center; margin-bottom: 32px;">
//         <h1 style="margin: 0; color: #0f172a; font-size: 28px; font-weight: 900; letter-spacing: -1px;">Who<span style="color: #533afd;">Posted</span></h1>
//     </div>
// `;
// const footerHtml = `
//     <div style="margin-top: 40px; padding-top: 24px; border-top: 1px solid #f1f5f9;">
//         <p style="margin: 0; color: #0f172a; font-weight: 700;">– Team WhoPosted</p>
//     </div>
//     <div style="margin-top: 40px; padding-top: 24px; border-top: 1px solid #f1f5f9; text-align: center;">
//         <p style="color: #cbd5e1; font-size: 11px; margin: 8px 0 0 0;">&copy; ${new Date().getFullYear()} WhoPosted AI. All rights reserved.</p>
//     </div>
// `;

// // 1. ACTIVE USERS EMAIL
// async function sendActiveUpdateEmail(to, clientName, jobs = [], trackingId = null) {
//     const htmlContent = `
//         <div style="${baseStyle}">
//             ${headerHtml}
//             <h2 style="font-weight: 900; font-size: 24px; color: #533afd; margin-bottom: 16px;">New Hiring Connections Added</h2>
//             <p style="font-size: 16px; line-height: 1.6;">Hi ${clientName},</p>
//             <p style="font-size: 16px; line-height: 1.6;">New job postings and hiring connections have been added to your WhoPosted dashboard.</p>
//             <p style="font-size: 16px; line-height: 1.6;">These are professionals actively hiring — giving you a chance to connect early and stand out before others apply.</p>
//             <p style="font-size: 16px; line-height: 1.6;">We've found 3 recruiters that look like a great fit for you today:</p>            
//             ${generateJobsHtml(jobs)}
//             <p style="font-size: 16px; line-height: 1.6; font-weight: 600;">Take a few minutes today to reach out and build meaningful connections.</p>
//             <div style="background: #533afd; padding: 24px; border-radius: 16px; text-align: center; margin: 32px 0;">
//                 <a href="https://www.whopostedai.com" style="display: inline-block; padding: 16px 32px; background: #ffffff; color: #533afd !important; text-decoration: none; border-radius: 12px; font-weight: 800; font-size: 14px; text-transform: uppercase;">Go to Dashboard</a>
//             </div>
//             ${footerHtml}
//             ${getTrackingPixelHtml(trackingId)}
//         </div>
//     `;
//     await sendMail(to, "You've got new job posted members", htmlContent);
// }

// // 2. EXPIRED USERS EMAIL
// async function sendRenewalEmail(to, clientName, jobs = [], trackingId = null) {
//     const htmlContent = `
//         <div style="${baseStyle}">
//             ${headerHtml}
//             <h2 style="font-weight: 900; font-size: 24px; color: #fda23aff; margin-bottom: 16px;">Restart Your Access to Hiring Contacts</h2>
//             <p style="font-size: 16px; line-height: 1.6;">Hi ${clientName},</p>
//             <p style="font-size: 16px; line-height: 1.6;">Your WhoPosted subscription has expired.</p>
//             <p style="font-size: 16px; line-height: 1.6;">You’re currently missing access to direct hiring contacts and recent job postings where early outreach can make a difference.</p>
//             <p style="font-size: 16px; line-height: 1.6;">Here are 3 recruiters active in the last 24 hours:</p>
//             ${generateJobsHtml(jobs, true)}
//             <p style="font-size: 16px; line-height: 1.6; font-weight: 600;">Renew your access to reconnect with hiring professionals and stay ahead in your job search.</p>
//             <div style="background: #0f172a; padding: 24px; border-radius: 16px; text-align: center; margin: 32px 0;">
//                 <a href="https://www.whopostedai.com/billing" style="display: inline-block; padding: 16px 32px; background: #533afd; color: #ffffff !important; text-decoration: none; border-radius: 12px; font-weight: 800; font-size: 14px; text-transform: uppercase;">Renew Access Now</a>
//             </div>
//             ${footerHtml}
//             ${getTrackingPixelHtml(trackingId)}
//         </div>
//     `;
//     await sendMail(to, "Don't miss out! Your Whoposted subscription has ended", htmlContent);
// }

// // 3. FREE USERS EMAIL
// async function sendUpsellEmail(to, clientName, jobs = [], trackingId = null) {
//     const htmlContent = `
//         <div style="${baseStyle}">
//             ${headerHtml}
//             <h2 style="font-weight: 900; font-size: 24px; color: #0f172a; margin-bottom: 16px;">Connect Directly with Hiring Managers</h2>
//             <p style="font-size: 16px; line-height: 1.6;">Hi ${clientName},</p>
//             <p style="font-size: 16px; line-height: 1.6;">Most candidates rely only on applications — but direct connections often lead to faster responses.</p>
//             <p style="font-size: 16px; line-height: 1.6;">WhoPosted gives you access to real job postings along with the LinkedIn profiles of hiring professionals.</p>
//             <p style="font-size: 16px; line-height: 1.6;">Here are 3 recruiters active in the last 24 hours:</p>
//             ${generateJobsHtml(jobs, true)}
//             <p style="font-size: 16px; line-height: 1.6; font-weight: 600;">Upgrade to unlock full access and start connecting directly.</p>
//             <div style="background: #533afd; padding: 24px; border-radius: 16px; text-align: center; margin: 32px 0;">
//                 <a href="https://www.whopostedai.com/pricing" style="display: inline-block; padding: 16px 32px; background: #ffffff; color: #533afd !important; text-decoration: none; border-radius: 12px; font-weight: 800; font-size: 14px; text-transform: uppercase;">Upgrade to Unlock</a>
//             </div>
//             ${footerHtml}
//             ${getTrackingPixelHtml(trackingId)}
//         </div>
//     `;
//     await sendMail(to, "Connect Directly with Hiring Managers", htmlContent);
// }

// module.exports = { 
//     sendActiveUpdateEmail, 
//     sendRenewalEmail, 
//     sendUpsellEmail 
// };
