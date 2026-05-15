/* global Office */

Office.onReady(() => {

    Office.actions.associate(
        "generateAIReply",
        generateAIReply
    );

});

function generateAIReply(event) {

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
                    result.value || "";

                console.log(
                    "Original Email:",
                    emailBody
                );

                const aiReply =
`
Hello,

Thank you for your email regarding the POC status.

We are currently reviewing the request and will provide an update shortly.

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