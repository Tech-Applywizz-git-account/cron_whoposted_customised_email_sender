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

export async function sendOTPEmail(to: string, otp: string) {
  try {
    const accessToken = await getMicrosoftAccessToken();
    const senderEmail = process.env.MS_SENDER_EMAIL || "krish@whopostedai.com";

    const htmlContent = `
      <div style="font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; max-width: 500px; margin: 0 auto; padding: 40px; border: 1px solid #f1f5f9; border-radius: 24px;">
        <div style="text-align: center; margin-bottom: 32px;">
          <h1 style="margin: 0; color: #0f172a; font-size: 28px; font-weight: 900; letter-spacing: -1px;">Who<span style="color: #533afd;">Posted</span></h1>
        </div>
        <h2 style="font-weight: 900; color: #533afd; margin-bottom: 24px;">Verify your email</h2>
        <p style="color: #64748b; font-size: 16px;">Welcome to WhoPosted! Use the verification code below to continue your signup.</p>
        <div style="background: #f8fafc; padding: 32px; border-radius: 16px; text-align: center; margin: 32px 0;">
          <span style="font-size: 40px; font-weight: 900; letter-spacing: 0.2em; color: #0f172a;">${otp}</span>
        </div>
        <p style="color: #94a3b8; font-size: 12px;">This code will expire in 10 minutes. If you didn't request this, you can safely ignore this email.</p>
        <div style="margin-top: 40px; padding-top: 24px; border-top: 1px solid #f1f5f9; text-align: center;">
          <p style="color: #94a3b8; font-size: 13px; margin: 0;">This email was sent by <strong>WhoPosted</strong>.</p>
          <p style="color: #cbd5e1; font-size: 12px; margin: 8px 0 0 0;">&copy; ${new Date().getFullYear()} WhoPosted AI. All rights reserved.</p>
        </div>
      </div>
    `;

    await fetch(`https://graph.microsoft.com/v1.0/users/${senderEmail}/sendMail`, {
      method: "POST",
      headers: { Authorization: `Bearer ${accessToken}`, "Content-Type": "application/json" },
      body: JSON.stringify({
        message: {
          subject: `${otp} is your WhoPosted verification code`,
          body: { contentType: "Html", content: htmlContent },
          toRecipients: [{ emailAddress: { address: to } }],
        }
      }),
    });
  } catch (err) {
    console.error("Failed to send OTP email:", err);
  }
}

export async function sendWelcomeEmail(to: string, fullName: string, password: string) {
  try {
    const accessToken = await getMicrosoftAccessToken();
    const senderEmail = process.env.MS_SENDER_EMAIL || "krish@whopostedai.com";

    const htmlContent = `
      <div style="font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; max-width: 550px; margin: 0 auto; padding: 40px; border: 1px solid #f1f5f9; border-radius: 24px;">
        <div style="text-align: center; margin-bottom: 32px;">
          <h1 style="margin: 0; color: #0f172a; font-size: 28px; font-weight: 900; letter-spacing: -1px;">Who<span style="color: #533afd;">Posted</span></h1>
        </div>
        <h2 style="font-weight: 900; color: #533afd; margin-bottom: 24px;">Welcome to WhoPosted, ${fullName}!</h2>
        <p style="color: #1e293b; font-size: 16px;">Your professional account has been created successfully. Here are your login credentials:</p>
        
        <div style="background: #f8fafc; padding: 24px; border-radius: 16px; margin: 24px 0;">
          <div style="display: flex; justify-content: space-between; margin-bottom: 12px;">
            <span style="color: #64748b; font-weight: 600;">Email</span>
            <span style="color: #0f172a; font-weight: 700;">${to}</span>
          </div>
          <div style="display: flex; justify-content: space-between;">
            <span style="color: #64748b; font-weight: 600;">Password</span>
            <span style="color: #0f172a; font-weight: 700;">${password}</span>
          </div>
        </div>

        <div style="text-align: center; margin-top: 32px;">
          <a href="https://www.whopostedai.com/" style="display: inline-block; padding: 16px 32px; background: #0f172a; color: #ffffff !important; text-decoration: none; border-radius: 12px; font-weight: 800; font-size: 14px; text-transform: uppercase;">Login to Dashboard</a>
        </div>
        <div style="margin-top: 40px; padding-top: 24px; border-top: 1px solid #f1f5f9; text-align: center;">
          <p style="color: #94a3b8; font-size: 13px; margin: 0;">This email was sent by <strong>WhoPosted</strong>.</p>
          <p style="color: #cbd5e1; font-size: 12px; margin: 8px 0 0 0;">&copy; ${new Date().getFullYear()} WhoPosted AI. All rights reserved.</p>
        </div>
      </div>
    `;

    await fetch(`https://graph.microsoft.com/v1.0/users/${senderEmail}/sendMail`, {
      method: "POST",
      headers: { Authorization: `Bearer ${accessToken}`, "Content-Type": "application/json" },
      body: JSON.stringify({
        message: {
          subject: "Welcome to WhoPosted! Your Account Credentials",
          body: { contentType: "Html", content: htmlContent },
          toRecipients: [{ emailAddress: { address: to } }],
        }
      }),
    });
  } catch (err) {
    console.error("Failed to send Welcome email:", err);
  }
}

export async function sendPaymentEmail({
  to,
  userName,
  transactionId,
  amount,
  paymentMethod,
  startDate,
  endDate,
}: {
  to: string;
  userName: string;
  transactionId: string;
  amount: string;
  paymentMethod: string;
  startDate: string;
  endDate: string;
}) {
  try {
    const accessToken = await getMicrosoftAccessToken();
    const senderEmail = process.env.MS_SENDER_EMAIL || "krish@whopostedai.com";

    const htmlContent = `
      <div style="font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; max-width: 600px; margin: 0 auto; padding: 40px; border: 1px solid #f1f5f9; border-radius: 24px;">
        <div style="text-align: center; margin-bottom: 32px;">
          <h1 style="margin: 0; color: #0f172a; font-size: 28px; font-weight: 900; letter-spacing: -1px;">Who<span style="color: #533afd;">Posted</span></h1>
        </div>
        <h2 style="font-weight: 900; font-size: 24px; color: #533afd; margin-bottom: 16px;">Professional Plan Upgrade Confirmed</h2>
        <p>Hi ${userName},</p>
        <p>Thank you for your payment. Your **WhoPosted Premium** subscription is now active.</p>
        
        <div style="background: #f8fafc; padding: 24px; border-radius: 16px; margin: 24px 0;">
          <div style="display: flex; justify-content: space-between; margin-bottom: 12px;">
            <span style="color: #64748b; font-weight: 600;">Transaction ID</span>
            <span style="color: #0f172a; font-weight: 700;">${transactionId}</span>
          </div>
          <div style="display: flex; justify-content: space-between; margin-bottom: 12px;">
            <span style="color: #64748b; font-weight: 600;">Amount Paid</span>
            <span style="color: #0f172a; font-weight: 700;">$${amount} USD</span>
          </div>
          <div style="display: flex; justify-content: space-between; margin-bottom: 12px;">
            <span style="color: #64748b; font-weight: 600;">Payment Method</span>
            <span style="color: #0f172a; font-weight: 700; text-transform: capitalize;">${paymentMethod}</span>
          </div>
          <div style="display: flex; justify-content: space-between; margin-bottom: 12px;">
            <span style="color: #64748b; font-weight: 600;">Plan Start Date</span>
            <span style="color: #0f172a; font-weight: 700;">${new Date(startDate).toLocaleDateString()}</span>
          </div>
          <div style="display: flex; justify-content: space-between;">
            <span style="color: #64748b; font-weight: 600;">Plan End Date</span>
            <span style="color: #533afd; font-weight: 700;">${new Date(endDate).toLocaleDateString()}</span>
          </div>
        </div>

        <div style="text-align: center; margin-top: 32px;">
          <a href="https://www.whopostedai.com/overview" style="display: inline-block; padding: 16px 32px; background: #0f172a; color: #ffffff !important; text-decoration: none; border-radius: 12px; font-weight: 800; font-size: 14px; text-transform: uppercase;">Go to Dashboard</a>
        </div>
        <div style="margin-top: 40px; padding-top: 24px; border-top: 1px solid #f1f5f9; text-align: center;">
          <p style="color: #94a3b8; font-size: 13px; margin: 0;">This email was sent by <strong>WhoPosted</strong>.</p>
          <p style="color: #cbd5e1; font-size: 12px; margin: 8px 0 0 0;">&copy; ${new Date().getFullYear()} WhoPosted AI. All rights reserved.</p>
        </div>
      </div>
    `;

    await fetch(`https://graph.microsoft.com/v1.0/users/${senderEmail}/sendMail`, {
      method: "POST",
      headers: { Authorization: `Bearer ${accessToken}`, "Content-Type": "application/json" },
      body: JSON.stringify({
        message: {
          subject: "WhoPosted Premium Upgrade - Payment Receipt",
          body: { contentType: "Html", content: htmlContent },
          toRecipients: [{ emailAddress: { address: to } }],
        }
      }),
    });
  } catch (error) {
    console.error("Failed to send payment email:", error);
  }
}

export async function sendSupportTicketEmail({
  to,
  clientEmail,
  clientName,
  ticketId,
  category,
  subject,
  description
}: {
  to: string;
  clientEmail: string;
  clientName: string;
  ticketId: string;
  category: string;
  subject: string;
  description: string;
}) {
  try {
    const accessToken = await getMicrosoftAccessToken();
    const senderEmail = process.env.MS_SENDER_EMAIL || "krish@whopostedai.com";

    const htmlContent = `
      <div style="font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; max-width: 600px; margin: 0 auto; padding: 40px; border: 1px solid #f1f5f9; border-radius: 24px;">
        <div style="text-align: center; margin-bottom: 32px;">
          <h1 style="margin: 0; color: #0f172a; font-size: 28px; font-weight: 900; letter-spacing: -1px;">Who<span style="color: #533afd;">Posted</span></h1>
        </div>
        <h2 style="font-weight: 900; font-size: 24px; color: #533afd; margin-bottom: 16px;">New Support Ticket Raised</h2>
        <p>A new ticket has been raised by a client.</p>
        
        <div style="background: #f8fafc; padding: 24px; border-radius: 16px; margin: 24px 0;">
          <div style="display: flex; justify-content: space-between; margin-bottom: 12px;">
            <span style="color: #64748b; font-weight: 600;">Client Name</span>
            <span style="color: #0f172a; font-weight: 700;">${clientName}</span>
          </div>
          <div style="display: flex; justify-content: space-between; margin-bottom: 12px;">
            <span style="color: #64748b; font-weight: 600;">Client Email</span>
            <span style="color: #0f172a; font-weight: 700;">${clientEmail}</span>
          </div>
          <div style="display: flex; justify-content: space-between; margin-bottom: 12px;">
            <span style="color: #64748b; font-weight: 600;">Category</span>
            <span style="color: #0f172a; font-weight: 700; text-transform: capitalize;">${category}</span>
          </div>
          <div style="display: flex; justify-content: space-between; margin-bottom: 12px;">
            <span style="color: #64748b; font-weight: 600;">Subject</span>
            <span style="color: #0f172a; font-weight: 700;">${subject}</span>
          </div>
          <div style="margin-top: 16px; border-top: 1px solid #e2e8f0; pt: 16px;">
            <span style="color: #64748b; font-weight: 600; display: block; margin-bottom: 8px;">Description</span>
            <p style="color: #1e293b; line-height: 1.6; margin: 0;">${description}</p>
          </div>
        </div>

        <div style="text-align: center; margin-top: 32px;">
          <a href="https://www.whopostedai.com/admin/tickets/${ticketId}" style="display: inline-block; padding: 16px 32px; background: #0f172a; color: #ffffff !important; text-decoration: none; border-radius: 12px; font-weight: 800; font-size: 14px; text-transform: uppercase;">View Ticket in Dashboard</a>
        </div>
        <div style="margin-top: 40px; padding-top: 24px; border-top: 1px solid #f1f5f9; text-align: center;">
          <p style="color: #94a3b8; font-size: 13px; margin: 0;">This email was sent by <strong>WhoPosted</strong>.</p>
          <p style="color: #cbd5e1; font-size: 12px; margin: 8px 0 0 0;">&copy; ${new Date().getFullYear()} WhoPosted AI. All rights reserved.</p>
        </div>
      </div>
    `;

    await fetch(`https://graph.microsoft.com/v1.0/users/${senderEmail}/sendMail`, {
      method: "POST",
      headers: { Authorization: `Bearer ${accessToken}`, "Content-Type": "application/json" },
      body: JSON.stringify({
        message: {
          subject: `New Support Ticket: ${subject} (${category})`,
          body: { contentType: "Html", content: htmlContent },
          toRecipients: [{ emailAddress: { address: to } }],
          replyTo: [{ emailAddress: { address: clientEmail } }]
        }
      }),
    });
  } catch (error) {
    console.error("Failed to send support ticket email:", error);
  }
}

export async function sendTicketReplyEmail({
  to,
  ticketId,
  subject,
  message,
  senderName,
  isAdminReply
}: {
  to: string;
  ticketId: string;
  subject: string;
  message: string;
  senderName: string;
  isAdminReply: boolean;
}) {
  try {
    const accessToken = await getMicrosoftAccessToken();
    const senderEmail = process.env.MS_SENDER_EMAIL || "krish@whopostedai.com";

    const dashboardLink = isAdminReply
      ? `https://www.whopostedai.com/tickets/${ticketId}`
      : `https://www.whopostedai.com/admin/tickets/${ticketId}`;

    const htmlContent = `
      <div style="font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; max-width: 600px; margin: 0 auto; padding: 40px; border: 1px solid #f1f5f9; border-radius: 24px;">
        <div style="text-align: center; margin-bottom: 32px;">
          <h1 style="margin: 0; color: #0f172a; font-size: 28px; font-weight: 900; letter-spacing: -1px;">Who<span style="color: #533afd;">Posted</span></h1>
        </div>
        <h2 style="font-weight: 900; font-size: 24px; color: #533afd; margin-bottom: 16px;">New Reply on Ticket</h2>
        <p>You have received a new message regarding ticket: <strong>${subject}</strong></p>
        
        <div style="background: #f8fafc; padding: 24px; border-radius: 16px; margin: 24px 0;">
          <p style="color: #64748b; font-weight: 600; margin-bottom: 8px; font-size: 12px; uppercase; tracking-widest;">From: ${senderName}</p>
          <div style="color: #1e293b; line-height: 1.6; font-size: 15px;">
            ${message}
          </div>
        </div>

        <div style="text-align: center; margin-top: 32px;">
          <a href="${dashboardLink}" style="display: inline-block; padding: 16px 32px; background: #0f172a; color: #ffffff !important; text-decoration: none; border-radius: 12px; font-weight: 800; font-size: 14px; text-transform: uppercase;">View Conversation</a>
        </div>
        <div style="margin-top: 40px; padding-top: 24px; border-top: 1px solid #f1f5f9; text-align: center;">
          <p style="color: #94a3b8; font-size: 13px; margin: 0;">This email was sent by <strong>WhoPosted</strong>.</p>
          <p style="color: #cbd5e1; font-size: 12px; margin: 8px 0 0 0;">&copy; ${new Date().getFullYear()} WhoPosted AI. All rights reserved.</p>
        </div>
      </div>
    `;

    await fetch(`https://graph.microsoft.com/v1.0/users/${senderEmail}/sendMail`, {
      method: "POST",
      headers: { Authorization: `Bearer ${accessToken}`, "Content-Type": "application/json" },
      body: JSON.stringify({
        message: {
          subject: `New Reply: ${subject}`,
          body: { contentType: "Html", content: htmlContent },
          toRecipients: [{ emailAddress: { address: to } }]
        }
      }),
    });
  } catch (error) {
    console.error("Failed to send ticket reply email:", error);
  }
}

export async function sendContactFormEmail({
  name,
  email,
  mobile,
  description
}: {
  name: string;
  email: string;
  mobile: string;
  description: string;
}) {
  try {
    const accessToken = await getMicrosoftAccessToken();
    const senderEmail = process.env.MS_SENDER_EMAIL || "krish@whopostedai.com";

    const htmlContent = `
      <div style="font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; max-width: 600px; margin: 0 auto; padding: 40px; border: 1px solid #f1f5f9; border-radius: 24px;">
        <div style="text-align: center; margin-bottom: 32px;">
          <h1 style="margin: 0; color: #0f172a; font-size: 28px; font-weight: 900; letter-spacing: -1px;">Who<span style="color: #533afd;">Posted</span></h1>
        </div>
        <h2 style="font-weight: 900; font-size: 24px; color: #533afd; margin-bottom: 16px;">New Inquiry from Landing Page</h2>
        <p>A visitor has submitted the contact form on the WhoPosted landing page.</p>
        
        <div style="background: #f8fafc; padding: 24px; border-radius: 16px; margin: 24px 0;">
          <div style="display: flex; justify-content: space-between; margin-bottom: 12px;">
            <span style="color: #64748b; font-weight: 600;">Name</span>
            <span style="color: #0f172a; font-weight: 700;">${name}</span>
          </div>
          <div style="display: flex; justify-content: space-between; margin-bottom: 12px;">
            <span style="color: #64748b; font-weight: 600;">Email</span>
            <span style="color: #0f172a; font-weight: 700;">${email}</span>
          </div>
          <div style="display: flex; justify-content: space-between; margin-bottom: 12px;">
            <span style="color: #64748b; font-weight: 600;">Mobile</span>
            <span style="color: #0f172a; font-weight: 700;">${mobile}</span>
          </div>
          <div style="margin-top: 16px; border-top: 1px solid #e2e8f0; padding-top: 16px;">
            <span style="color: #64748b; font-weight: 600; display: block; margin-bottom: 8px;">Description/Message</span>
            <p style="color: #1e293b; line-height: 1.6; margin: 0;">${description}</p>
          </div>
        </div>
        <div style="margin-top: 40px; padding-top: 24px; border-top: 1px solid #f1f5f9; text-align: center;">
          <p style="color: #94a3b8; font-size: 13px; margin: 0;">This email was sent by <strong>WhoPosted</strong>.</p>
          <p style="color: #cbd5e1; font-size: 12px; margin: 8px 0 0 0;">&copy; ${new Date().getFullYear()} WhoPosted AI. All rights reserved.</p>
        </div>
      </div>
    `;

    await fetch(`https://graph.microsoft.com/v1.0/users/${senderEmail}/sendMail`, {
      method: "POST",
      headers: { Authorization: `Bearer ${accessToken}`, "Content-Type": "application/json" },
      body: JSON.stringify({
        message: {
          subject: `New Landing Page Inquiry: ${name}`,
          body: { contentType: "Html", content: htmlContent },
          toRecipients: [{ emailAddress: { address: "krish@whopostedai.com" } }],
          replyTo: [{ emailAddress: { address: email } }]
        }
      }),
    });
  } catch (error) {
    console.error("Failed to send contact form email:", error);
    throw error;
  }
}

export async function sendCancellationEmail({
  to,
  userName,
  endDate
}: {
  to: string;
  userName: string;
  endDate: string;
}) {
  try {
    const accessToken = await getMicrosoftAccessToken();
    const senderEmail = process.env.MS_SENDER_EMAIL || "krish@whopostedai.com";

    const htmlContent = `
      <div style="font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; max-width: 600px; margin: 0 auto; padding: 40px; border: 1px solid #f1f5f9; border-radius: 24px;">
        <div style="text-align: center; margin-bottom: 32px;">
          <h1 style="margin: 0; color: #0f172a; font-size: 28px; font-weight: 900; letter-spacing: -1px;">Who<span style="color: #533afd;">Posted</span></h1>
        </div>
        <h2 style="font-weight: 900; font-size: 24px; color: #f43f5e; margin-bottom: 16px;">Subscription Cancelled</h2>
        <p>Hi ${userName},</p>
        <p>We're sorry to see you go. Your subscription has been cancelled as per your request.</p>
        
        <div style="background: #fff1f2; padding: 24px; border-radius: 16px; margin: 24px 0; border: 1px solid #ffe4e6;">
          <p style="color: #9f1239; font-weight: 700; margin-bottom: 8px;">Service Status</p>
          <p style="color: #1e293b; margin: 0;">
            Your premium services will remain active until <strong>${new Date(endDate).toLocaleDateString()}</strong>.
          </p>
        </div>

        <p style="color: #64748b; font-size: 14px; line-height: 1.6;">
          To prevent any interruption in your professional job search insights, you can reactivate your subscription at any time through your dashboard.
        </p>

        <div style="text-align: center; margin-top: 32px;">
          <a href="https://www.whopostedai.com/billing" style="display: inline-block; padding: 16px 32px; background: #0f172a; color: #ffffff !important; text-decoration: none; border-radius: 12px; font-weight: 800; font-size: 14px; text-transform: uppercase;">Renew Subscription</a>
        </div>
        
        <p style="margin-top: 32px; color: #94a3b8; font-size: 12px; text-align: center;">
          If you have any questions, feel free to contact us at krish@whopostedai.com
        </p>
        <div style="margin-top: 40px; padding-top: 24px; border-top: 1px solid #f1f5f9; text-align: center;">
          <p style="color: #94a3b8; font-size: 13px; margin: 0;">This email was sent by <strong>WhoPosted</strong>.</p>
          <p style="color: #cbd5e1; font-size: 12px; margin: 8px 0 0 0;">&copy; ${new Date().getFullYear()} WhoPosted AI. All rights reserved.</p>
        </div>
      </div>
    `;

    await fetch(`https://graph.microsoft.com/v1.0/users/${senderEmail}/sendMail`, {
      method: "POST",
      headers: { Authorization: `Bearer ${accessToken}`, "Content-Type": "application/json" },
      body: JSON.stringify({
        message: {
          subject: "Your WhoPosted Subscription has been Cancelled",
          body: { contentType: "Html", content: htmlContent },
          toRecipients: [{ emailAddress: { address: to } }]
        }
      }),
    });
  } catch (error) {
    console.error("Failed to send cancellation email:", error);
  }
}
