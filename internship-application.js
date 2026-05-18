function getSelectedInternshipAreas() {
  return Array.from(document.querySelectorAll('input[name="internshipArea"]:checked')).map(
    (checkbox) => checkbox.value
  );
}

function setInternshipAreaError(show) {
  const options = document.getElementById("internshipAreaOptions");
  const error = document.getElementById("internshipAreaError");
  options.classList.toggle("is-invalid", show);
  error.hidden = !show;
}

function setSubmitSpinner(show) {
  const spinner = document.getElementById("submitSpinner");
  const submitButton = document.getElementById("submitApplicationBtn");
  spinner.hidden = !show;
  submitButton.disabled = show;
}

document.querySelectorAll('input[name="internshipArea"]').forEach((checkbox) => {
  checkbox.addEventListener("change", () => {
    if (getSelectedInternshipAreas().length > 0) {
      setInternshipAreaError(false);
    }
  });
});

document.getElementById("internshipApplicationForm").addEventListener("submit", async function (e) {
  e.preventDefault();

  const form = document.getElementById("internshipApplicationForm");
  const thankYouScreen = document.getElementById("thankYouScreen");

  const internshipAreas = getSelectedInternshipAreas();
  if (internshipAreas.length === 0) {
    setInternshipAreaError(true);
    document.getElementById("internshipAreaGroup").scrollIntoView({ behavior: "smooth", block: "center" });
    return;
  }

  const flowUrl =
    "https://defaultdf0fc509acb44023b500fdf382dde4.30.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/647f954dcd68454ea4e0b11ac3bd37a6/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=iu3hOjM9jFItmu5-q2s4dKKAV0zg0K3o3h8HgIBbito";

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

  setSubmitSpinner(true);

  try {
    const payload = {
      fullName: document.getElementById("fullName").value.trim(),
      email: document.getElementById("email").value.trim(),
      studentId: document.getElementById("studentId").value.trim(),
      faculty: document.getElementById("faculty").value.trim(),
      gpa: document.getElementById("gpa").value,
      creditsCompleted: document.getElementById("creditsCompleted").value,
      degreeDuration: document.getElementById("degreeDuration").value.trim(),
      phone: document.getElementById("phone").value.trim(),
      willingTwoMonths: document.getElementById("willingTwoMonths").value,
      internshipArea: internshipAreas,
      transcript: await fileToBase64("transcriptAttachment"),
      video: await fileToBase64("internshipVideo")
    };

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
    } else {
      console.error("Submission failed:", response.status, await response.text());
      alert("Something went wrong while submitting your application. Please try again.");
      setSubmitSpinner(false);
    }
  } catch (error) {
    console.error("Submission error:", error);
    alert("Submission failed. Please check your connection and try again.");
    setSubmitSpinner(false);
  }
});
