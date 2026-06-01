// --- BLOCK: Task Pane UI Logic ---
// Purpose: Handle UI interactions, form rendering, and orchestration of save/update operations

let currentContact = null;
let existingContact = null;
let excelFileId = null;

Office.onReady(async (info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("btn-save").addEventListener("click", handleSave);
    document.getElementById("btn-update").addEventListener("click", handleUpdate);
    document.getElementById("btn-cancel").addEventListener("click", handleCancel);
    document.getElementById("btn-login").addEventListener("click", handleLogin);

    await Auth.initMsal();

    if (Auth.isLoggedIn()) {
      await loadEmailData();
    } else {
      showLoginScreen();
    }
  }
});

async function loadEmailData() {
  const item = Office.context.mailbox.item;
  const bodyPromise = new Promise((resolve, reject) => {
    item.body.getAsync(Office.CoercionType.Text, { asyncContext: "body" }, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(result.value);
      } else {
        reject(new Error(result.error.message));
      }
    });
  });

  const body = await bodyPromise;

  const senderEmail = item.from && item.from.emailAddress ? item.from.emailAddress : "";
  const senderName = item.from && item.from.displayName ? item.from.displayName : "";

  currentContact = Extractor.extractFromEmail(body, senderEmail, senderName);
  renderForm(currentContact);

  const token = await Auth.getToken();
  existingContact = await Graph.searchContactByEmail(currentContact.email, token);

  if (existingContact) {
    showDuplicateWarning(existingContact);
  }

  excelFileId = await Graph.findOrCreateExcelFile(token);
  await Graph.ensureHeaders(excelFileId, token);
}

function renderForm(contact) {
  const fields = ["firstName", "lastName", "email", "mobile", "phone", "title", "company", "street", "city", "postal", "country"];
  fields.forEach((key) => {
    const el = document.getElementById("field-" + key);
    if (el) {
      el.value = contact[key] || "";
    }
  });
}

function getFormData() {
  const fields = ["firstName", "lastName", "email", "mobile", "phone", "title", "company", "street", "city", "postal", "country"];
  const contact = {};
  fields.forEach((key) => {
    const el = document.getElementById("field-" + key);
    contact[key] = el ? el.value : "";
  });
  return contact;
}

function showDuplicateWarning(existing) {
  const displayName = existing.displayName || (existing.givenName + " " + existing.surname);
  document.getElementById("duplicate-text").textContent = "Contact already exists: " + displayName;
  document.getElementById("duplicate-banner").style.display = "block";
  document.getElementById("btn-save").style.display = "none";
  document.getElementById("btn-update").style.display = "inline-block";
  document.getElementById("btn-cancel").style.display = "inline-block";
}

function hideDuplicateWarning() {
  document.getElementById("duplicate-banner").style.display = "none";
  document.getElementById("btn-save").style.display = "inline-block";
  document.getElementById("btn-update").style.display = "none";
}

async function handleSave() {
  const contact = getFormData();
  if (!contact.email) {
    showStatus("Email is required", "error");
    return;
  }

  const btn = document.getElementById("btn-save");
  btn.disabled = true;
  btn.textContent = "Saving...";

  try {
    const token = await Auth.getToken();
    await Graph.createContact(contact, token);
    try {
      await Graph.appendContactToExcel(excelFileId, contact, token);
      showStatus("Contact saved successfully", "success");
    } catch (e) {
      showStatus("Saved to contacts. Excel update failed: " + e.message, "warning");
    }
  } catch (e) {
    showStatus("Error saving contact: " + e.message, "error");
  } finally {
    btn.disabled = false;
    btn.textContent = "Save Contact";
  }
}

async function handleUpdate() {
  const contact = getFormData();
  if (!contact.email) {
    showStatus("Email is required", "error");
    return;
  }

  const btn = document.getElementById("btn-update");
  btn.disabled = true;
  btn.textContent = "Updating...";

  try {
    const token = await Auth.getToken();
    await Graph.updateContact(existingContact.id, contact, token);
    const row = await Graph.findExcelRow(excelFileId, contact.email, token);
    if (row) {
      await Graph.updateContactInExcel(excelFileId, row, contact, token);
    } else {
      await Graph.appendContactToExcel(excelFileId, contact, token);
    }
    showStatus("Contact updated successfully", "success");
  } catch (e) {
    showStatus("Error updating contact: " + e.message, "error");
  } finally {
    btn.disabled = false;
    btn.textContent = "Update Existing";
  }
}

function handleCancel() {
  const fields = ["firstName", "lastName", "email", "mobile", "phone", "title", "company", "street", "city", "postal", "country"];
  fields.forEach((key) => {
    const el = document.getElementById("field-" + key);
    if (el) {
      el.value = "";
    }
  });
  hideDuplicateWarning();
  existingContact = null;
}

function showLoginScreen() {
  document.getElementById("login-screen").style.display = "block";
  document.getElementById("main-content").style.display = "none";
}

async function handleLogin() {
  try {
    await Auth.signIn();
    document.getElementById("login-screen").style.display = "none";
    document.getElementById("main-content").style.display = "block";
    await loadEmailData();
  } catch (e) {
    showStatus("Login failed: " + e.message, "error");
  }
}

function showStatus(message, type) {
  const el = document.getElementById("status-message");
  el.textContent = message;
  el.className = "status-" + type;
  el.style.display = "block";
  setTimeout(() => {
    el.style.display = "none";
  }, 4000);
}