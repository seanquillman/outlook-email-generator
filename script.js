Office.onReady(() => {
  console.log("Add-in is ready");
});

const backendUrl = "http://127.0.0.1:5000/generate_email";

async function generateEmail() {
  const formality = document.getElementById("formality").value;
  const audience = document.getElementById("audience").value;
  const length = document.getElementById("length").
  const length = document.getElementById("length").value;
  const keyPoints = document.getElementById("key-points").value;

  const prompt = `Please write an ${formality} email to ${audience} that is approximately ${length} words long and covers the following points: ${keyPoints}`;

  const response = await fetch(backendUrl, {
    method: "POST",
    headers: {
      "Content-Type": "application/json"
    },
    body: JSON.stringify({ prompt })
  });

  const generatedEmail = await response.text();

  if (Office.context.mailbox.item != null) {
    Office.context.mailbox.item.body.setSelectedDataAsync(
      generatedEmail,
      { coercionType: Office.CoercionType.Html },
      function (asyncResult) {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          console.error("Failed to insert email content: " + asyncResult.error.message);
        } else {
          console.log("Email content inserted");
        }
      }
    );
  }
}
