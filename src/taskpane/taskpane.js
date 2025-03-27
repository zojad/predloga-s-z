Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    // Set up button handlers with error handling
    document.getElementById("checkTextButton").onclick = async () => {
      try {
        if (typeof window.checkDocumentText !== 'function') {
          throw new Error("Check document function not loaded");
        }
        await window.checkDocumentText();
      } catch (error) {
        showError("Failed to check document: " + error.message);
      }
    };

    document.getElementById("acceptChangeButton").onclick = async () => {
      try {
        if (typeof window.acceptCurrentChange !== 'function') {
          throw new Error("Accept function not loaded");
        }
        await window.acceptCurrentChange();
      } catch (error) {
        showError("Failed to accept change: " + error.message);
      }
    };

    document.getElementById("rejectChangeButton").onclick = async () => {
      try {
        if (typeof window.rejectCurrentChange !== 'function') {
          throw new Error("Reject function not loaded");
        }
        await window.rejectCurrentChange();
      } catch (error) {
        showError("Failed to reject change: " + error.message);
      }
    };

    document.getElementById("acceptAllButton").onclick = async () => {
      try {
        if (typeof window.acceptAllChanges !== 'function') {
          throw new Error("Accept-all function not loaded");
        }
        await window.acceptAllChanges();
      } catch (error) {
        showError("Failed to accept all changes: " + error.message);
      }
    };

    document.getElementById("rejectAllButton").onclick = async () => {
      try {
        if (typeof window.rejectAllChanges !== 'function') {
          throw new Error("Reject-all function not loaded");
        }
        await window.rejectAllChanges();
      } catch (error) {
        showError("Failed to reject all changes: " + error.message);
      }
    };
  }
});

// Error display function (only shows when errors occur)
function showError(message) {
  // Remove any existing error messages
  const oldError = document.getElementById("error-message");
  if (oldError) oldError.remove();

  // Create and display new error message
  const errorDiv = document.createElement("div");
  errorDiv.id = "error-message";
  errorDiv.style.color = "#a80000";
  errorDiv.style.padding = "10px";
  errorDiv.style.marginTop = "10px";
  errorDiv.style.border = "1px solid #a80000";
  errorDiv.style.borderRadius = "4px";
  errorDiv.style.backgroundColor = "#fde7e9";
  errorDiv.textContent = message;

  // Insert after the button container
  const container = document.querySelector(".button-container");
  container.insertAdjacentElement("afterend", errorDiv);

  // Auto-remove after 10 seconds
  setTimeout(() => {
    errorDiv.remove();
  }, 10000);
}
