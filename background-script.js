Office.onReady(() => {
  Excel.run(async (context) => {
    const worksheet = context.workbook.worksheets.getItem("Master Input");

    // Register the onChanged event
    worksheet.onChanged.add(async (eventArgs) => {
      console.log("Change detected:");
      console.log("Change Type:", eventArgs.changeType);
      console.log("Changed Range Address:", eventArgs.address);
    });

    await context.sync();
    console.log("Event handler registered successfully.");
  }).catch((error) => {
    console.error("Error initializing background script:", error);
  });
});
