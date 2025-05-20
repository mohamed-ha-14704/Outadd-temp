Office.onReady(() => {
  console.log("Office ready");
});

function main(event) {
  Office.context.mailbox.item.notificationMessages.addAsync("onsend-msg", {
    type: "informationalMessage",
    message: "Popup before sending email!",
    icon: "icon16",
    persistent: false
  });
  event.completed({ allowEvent: true }); // allow sending
}
function mainHandleAttachments(event) {
    try {
        const item = Office.context.mailbox.item;

        item.attachments.getAsync((result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                const attachments = result.value;
                const attachmentCount = attachments.length;

                // Show a message to the user
                Office.context.mailbox.item.notificationMessages.replaceAsync(
                    "attachmentNotice",
                    {
                        type: "informationalMessage",
                        message: `You now have ${attachmentCount} attachment(s) on this item.`,
                        icon: "icon16", // define this in your manifest if you want
                        persistent: false
                    },
                    (asyncResult) => {
                        if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
                            console.error("Failed to show notification:", asyncResult.error.message);
                        }
                        event.completed();
                    }
                );
            } else {
                console.error("Error fetching attachments:", result.error);
                event.completed();
            }
        });
    } catch (err) {
        console.error("Unhandled error:", err);
        event.completed();
    }
}
