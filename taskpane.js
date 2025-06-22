function applyLabel() {
  const classification = document.getElementById("classification-select").value;
  Office.onReady().then(() => {
    Word.run(async (context) => {
      const sections = context.document.body.sections;
      context.load(sections, "items");
      await context.sync();

      sections.items.forEach(section => {
        const header = section.getHeader("Primary");
        header.insertText(`[ ${classification} ]`, Word.InsertLocation.replace);
      });

      await context.sync();
    }).catch(error => {
      console.error("Error applying classification:", error);
    });
  });
}