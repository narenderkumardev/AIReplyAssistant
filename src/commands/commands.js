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

        item.body.getAsync(
            Office.CoercionType.Text,

            function(result) {

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
                    result.value;

                const aiReply =
`
Hello,

Thank you for your email.

We are reviewing your request and will get back to you shortly.

Regards,
Support Team

—
Generated with AI assistance.
`;

                item.displayReplyForm({

                    htmlBody:
                        aiReply.replace(
                            /\n/g,
                            "<br>"
                        )
                });

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