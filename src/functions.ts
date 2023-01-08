async function onSend(event: Office.AddinCommands.Event) {
  await Office.onReady();

  setTimeout(() => {
    Office.context.mailbox.item.loadCustomPropertiesAsync((result) => {
      if (result.status === Office.AsyncResultStatus.Failed) {
        const { name, code, message } = result.error ?? {};
        const failureMessage = `[Error] Name: ${name}, Code: ${code}, Message: ${message}`;

        Office.context.mailbox.item.body.setAsync(
          `Failed to load custom properties with following error ;(

${failureMessage}`,
          { coercionType: Office.CoercionType.Text },
          () => {
            event.completed({ allowEvent: false });
          }
        );
      } else {
        event.completed({ allowEvent: true });
      }
    });
  }, 3000);
}

Object.assign(globalThis, {
  onSend,
});

Office.onReady();
