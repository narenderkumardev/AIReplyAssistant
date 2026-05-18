/* global Office */

Office.onReady(() => {

    Office.actions.associate(
        "generateAIReply",
        generateAIReply
    );

});

async function generateAIReply(event) {

    const startTime =
        Date.now();

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

                let mailboxUser =
                    "unknown@local";

                if (
                    Office.context &&
                    Office.context.mailbox &&
                    Office.context.mailbox.userProfile
                ) {

                    mailboxUser =
                        Office.context.mailbox
                        .userProfile.emailAddress;
                }

                console.log(
                    "Mailbox User:",
                    mailboxUser
                );

                // LOAD LOCAL SETTINGS

                let settings =
                    JSON.parse(
                        localStorage.getItem(
                            "ai_reply_settings"
                        ) || "{}"
                    );

                const tone =
                    settings.tone ||
                    "Professional";

                // TEMP CONFIDENCE
                // Future Gemini/OpenAI

                const confidence =
                    85;

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

                    // RESPONSE TIME

                    const responseTime =
                        Math.floor(
                            (
                                Date.now()
                                - startTime
                            ) / 1000
                        );

                    // SAVE INSIGHTS
                    // BACKEND API PLACEHOLDER

                    try {

                        await fetch(

                            "https://your-api-url/api/insights",

                            {

                                method:
                                    "POST",

                                headers: {

                                    "Content-Type":
                                        "application/json"
                                },

                                body:
                                    JSON.stringify({

                                        mailbox_user:
                                            mailboxUser,

                                        source:
                                            "AI_REPLY",

                                        action_type:
                                            "DRAFT",

                                        confidence:
                                            confidence,

                                        response_time_seconds:
                                            responseTime
                                    })
                            }
                        );

                        console.log(
                            "Insights API called successfully."
                        );

                    }
                    catch (apiError) {

                        console.error(
                            "Insights API Error:",
                            apiError
                        );
                    }

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
