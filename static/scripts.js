document.addEventListener("DOMContentLoaded", () => {
  const form = document.getElementById("upload-form");
  const overlay = document.getElementById("loading-overlay");

  form.addEventListener("submit", async (event) => {
    event.preventDefault();

    overlay.style.display = "flex";
    document.body.classList.add("blur");

    const formData = new FormData(form);

    try {
      const response = await fetch(form.action, {
        method: form.method,
        body: formData,
      });

      if (response.ok) {
        const blob = await response.blob();
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement("a");
        a.style.display = "none";
        a.href = url;
        a.download = formData.get("file-name") + ".xlsx";
        document.body.appendChild(a);
        a.click();
        window.URL.revokeObjectURL(url);

        // Clear the form fields
        form.reset();
      } else {
        console.error("An error occurred during file processing.");
      }
    } catch (error) {
      console.error("An error occurred during form submission:", error);
    } finally {
      overlay.style.display = "none";
      document.body.classList.remove("blur");
    }
  });
});
