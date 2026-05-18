/* global Office */

const supabaseClient =
    window.supabase.createClient(

        "https://vpszsnlevsrphplciitx.supabase.co",

        "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InZwc3pzbmxldnNycGhwbGNpaXR4Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3Nzg4NDE5MTksImV4cCI6MjA5NDQxNzkxOX0.zHm_4_css3VEIDqWGAi6oUrsDI9PdE-FkY5mvKBAZfU"
    );

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

                const mailboxUser =
                    Office.context.mailbox
                    .userProfile.emailAddress;

                console.log(
                    "Mailbox User:",
                    mailboxUser
                );

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
                // Later Gemini/OpenAI will provide this

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

                    // SAVE INSIGHTS

                    const responseTime =
                        Math.floor(
                            (
                                Date.now()
                                - startTime
                            ) / 1000
                        );

                    const {
                        error
                    } =
                        await supabaseClient
                        .from(
                            "email_insights"
                        )
                        .insert({

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
                        });

                    if (error) {

                        console.error(
                            "Insights Save Error:",
                            error
                        );

                    } else {

                        console.log(
                            "Insights saved successfully."
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