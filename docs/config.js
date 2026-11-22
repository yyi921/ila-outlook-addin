/* iLA add-in settings. Nothing here is secret (an Entra client ID is public). */
window.ILA_CONFIG = {
  // Where the combined invoice + DDTA PDFs are emailed, one PDF per email.
  lightyearEmail: "theville@lightyear.cloud",

  // Entra app registration (Application/client ID) that lets the add-in send mail
  // as the signed-in user. Leave empty and the pane can still build and download PDFs.
  // Setup steps are in README.md, section "Let the add-in send email".
  clientId: "22e933ba-182d-467f-a318-8f1ac9d5a222",
};
