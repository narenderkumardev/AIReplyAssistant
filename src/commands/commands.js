/* global Office */

Office.onReady(() => {

    Office.actions.associate(
        "generateAIReply",
        generateAIReply
    );

});

async function generateAIReply(event) {

    try {

        const item =
            Office.context.mailbox.item;

        if (!item) {

            console.error(
                "No mailbox item found."
            );

            event.completed();

            return;
        }

        item.body.getAsync(
            Office.CoercionType.Text,

            async function(result) {

                if (
                    result.status !==
                    Office.AsyncResultStatus.Succeeded
                ) {

                    console.error(
                        "Failed to read email body."
                    );

                    event.completed();

                    return;
                }

                const emailBody =
                    result.value || "";

                console.log(
                    "Original Email:",
                    emailBody
                );

                const mailboxUser =
                    Office.context.mailbox.userProfile.emailAddress;

                console.log(
                    "Mailbox User:",
                    mailboxUser
                );

                // LOCAL CACHE

                let settings =
                    JSON.parse(
                        localStorage.getItem(
                            "ai_reply_settings"
                        ) || "{}"
                    );

                const tone =
                    settings.tone ||
                    "Professional";

                // FUTURE BACKEND FLOW
                // Here later you will:
                // 1. Call REST API
                // 2. Load settings from Supabase
                // 3. Generate Gemini/OpenAI response
                // 4. Save insights

                const aiReply =
`
Hello,

Thank you for your email.

We are reviewing your request and will get back to you shortly.

Response Tone:
${tone}

Regards,
Support Team

-----------------------------------
AI Generated Draft Response
Please review before sending.
-----------------------------------
`;

                try {

                    item.displayReplyForm(
                        aiReply
                    );

                    console.log(
                        "Reply form opened successfully."
                    );

                    // FUTURE:
                    // Save insights via REST API

                }
                catch (replyError) {

                    console.error(
                        "displayReplyForm error:",
                        replyError
                    );
                }

                event.completed();
            }
        );

    }
    catch (error) {

        console.error(
            "AI Reply Error:",
            error
        );

        event.completed();
    }
}