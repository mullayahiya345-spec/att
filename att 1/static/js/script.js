document.addEventListener("DOMContentLoaded", () => {
  const attendanceButtons = document.querySelectorAll(".attendance-btn");
  attendanceButtons.forEach((button) => {
    button.addEventListener("click", async () => {
      const subjectId = button.dataset.subjectId;
      const status = button.dataset.status;
      const formData = new FormData();
      formData.append("status", status);

      try {
        const response = await fetch(`/api/subjects/${subjectId}/attendance`, {
          method: "POST",
          body: formData,
        });
        const result = await response.json();
        if (result.success) {
          const percentElement = document.getElementById(`percent-${subjectId}`);
          if (percentElement) {
            percentElement.textContent = `${result.percent}%`;
          }
          const badgeElement = document.getElementById(`badge-${subjectId}`);
          if (badgeElement) {
            badgeElement.className = `badge rounded-pill badge-${result.badge_class}`;
            badgeElement.textContent = result.label;
          }
          const message = document.getElementById("attendanceMessage");
          if (message) {
            message.textContent = result.message;
          }
        } else {
          alert(result.message || "Unable to update attendance.");
        }
      } catch (error) {
        console.error(error);
        alert("Unable to update attendance right now.");
      }
    });
  });
});
