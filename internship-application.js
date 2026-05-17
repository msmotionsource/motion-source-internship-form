document.getElementById("internshipApplicationForm").addEventListener("submit", async function (e) {
  e.preventDefault();

  const form = document.getElementById("internshipApplicationForm");
  const thankYouScreen = document.getElementById("thankYouScreen");
  const submitButton = form.querySelector("button[type='submit']");

  const flowUrl =
    "https://1f8ef28f522de09a84be445609784c.0e.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/327e626637624b8eadd057629eabc5f9/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=CXmcV9EP1qfR_DZzBqeVAV9GHfa-1CB6JxTQy2aOSv8";

  async function fileToBase64(fileInputId) {
    const file = document.getElementById(fileInputId).files[0];
    if (!file) {
      throw new Error("Missing file: " + fileInputId);
    }

    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => {
        resolve({
          fileName: file.name,
          contentType: file.type,
          fileContent: reader.result.split(",")[1]
        });
      };
      reader.onerror = reject;
      reader.readAsDataURL(file);
    });
  }

  const payload = {
    fullName: document.getElementById("fullName").value.trim(),
    email: document.getElementById("email").value.trim(),
    studentId: document.getElementById("studentId").value.trim(),
    faculty: document.getElementById("faculty").value.trim(),
    gpa: document.getElementById("gpa").value,
    phone: document.getElementById("phone").value.trim(),
    willingTwoMonths: document.getElementById("willingTwoMonths").value,
    internshipArea: document.getElementById("internshipArea").value,
    transcript: await fileToBase64("transcriptAttachment"),
    video: await fileToBase64("internshipVideo")
  };

  submitButton.disabled = true;
  submitButton.textContent = "Submitting...";

  try {
    const response = await fetch(flowUrl, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload)
    });

    if (response.ok) {
      form.style.display = "none";
      thankYouScreen.style.display = "block";
      document.getElementById("formIntroText").style.display = "none";
      window.scrollTo({ top: 0, behavior: "smooth" });
      // window.location.href = "/thank-you";
    } else {
      alert("Something went wrong while submitting your application. Please try again.");
      submitButton.disabled = false;
      submitButton.textContent = "Submit Application";
    }
  } catch (error) {
    alert("Submission failed. Please check your connection and try again.");
    submitButton.disabled = false;
    submitButton.textContent = "Submit Application";
  }
});
