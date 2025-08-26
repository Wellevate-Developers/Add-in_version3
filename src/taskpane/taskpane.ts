/* global document, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.PowerPoint) {
    document.getElementById("sideload-msg")!.style.display = "none";
    document.getElementById("app-body")!.style.display = "flex";
    document.getElementById("run")!.onclick = run;
  }
});

export async function run() {
  try {
    await PowerPoint.run(async (context) => {
      // Get the first slide
      const slide = context.presentation.slides.getItemAt(0);

      // Add a text box
      slide.shapes.addTextBox("Hello from Office Add-in!", {
        left: 100,
        top: 100,
        width: 300,
        height: 50,
      });

      await context.sync();
    });
  } catch (error) {
    console.error(error);
  }
}
