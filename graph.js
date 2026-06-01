// --- BLOCK: Graph API Helper Functions ---
// Purpose: Provide functions to interact with Microsoft Graph API for contacts and Excel files

const GRAPH_BASE = "https://graph.microsoft.com/v1.0";

async function graphRequest(method, path, token, body = null) {
  const url = GRAPH_BASE + path;
  const headers = {
    "Authorization": "Bearer " + token,
    "Content-Type": "application/json"
  };
  const options = {
    method: method,
    headers: headers
  };
  if (body) {
    options.body = JSON.stringify(body);
  }
  const response = await fetch(url, options);
  if (response.status === 204) {
    return null;
  }
  if (!response.ok) {
    const text = await response.text();
    throw new Error("Graph API error " + response.status + ": " + text);
  }
  return response.json();
}

async function searchContactByEmail(email, token) {
  const encodedEmail = encodeURIComponent(email);
  const path = "/me/contacts?$filter=emailAddresses/any(e:e/address eq '" + encodedEmail + "')";
  const data = await graphRequest("GET", path, token);
  if (data && data.value && data.value.length > 0) {
    return data.value[0];
  }
  return null;
}

async function createContact(contact, token) {
  const body = {
    givenName: contact.firstName,
    surname: contact.lastName,
    emailAddresses: [{ address: contact.email, name: contact.firstName + " " + contact.lastName }],
    mobilePhone: contact.mobile || "",
    businessPhones: contact.phone ? [contact.phone] : [],
    jobTitle: contact.title || "",
    companyName: contact.company || "",
    businessAddress: {
      street: contact.street || "",
      city: contact.city || "",
      postalCode: contact.postal || "",
      countryOrRegion: contact.country || ""
    }
  };
  return graphRequest("POST", "/me/contacts", token, body);
}

async function updateContact(id, contact, token) {
  const body = {
    givenName: contact.firstName,
    surname: contact.lastName,
    emailAddresses: [{ address: contact.email, name: contact.firstName + " " + contact.lastName }],
    mobilePhone: contact.mobile || "",
    businessPhones: contact.phone ? [contact.phone] : [],
    jobTitle: contact.title || "",
    companyName: contact.company || "",
    businessAddress: {
      street: contact.street || "",
      city: contact.city || "",
      postalCode: contact.postal || "",
      countryOrRegion: contact.country || ""
    }
  };
  return graphRequest("PATCH", "/me/contacts/" + id, token, body);
}

async function findOrCreateExcelFile(token) {
  try {
    const data = await graphRequest("GET", "/me/drive/root:/OutlookContacts.xlsx", token);
    return data.id;
  } catch (e) {
    if (e.message && e.message.includes("404")) {
      const createBody = {
        "@microsoft.graph.conflict": "replace"
      };
      const response = await fetch(GRAPH_BASE + "/me/drive/root:/OutlookContacts.xlsx:/content", {
        method: "PUT",
        headers: {
          "Authorization": "Bearer " + token,
          "Content-Type": "text/plain"
        },
        body: ""
      });
      const data = await response.json();
      return data.id;
    }
    throw e;
  }
}

async function ensureHeaders(fileId, token) {
  try {
    const body = {
      values: [["First Name", "Last Name", "Email", "Mobile", "Phone", "Title", "Company", "Street", "City", "Postal Code", "Country", "Date Added", "Date Updated"]]
    };
    await graphRequest("PATCH", "/me/drive/items/" + fileId + "/workbook/worksheets('Sheet1')/range(address='A1:M1')", token, body);
  } catch (e) {
    // Non-fatal if headers already exist
  }
}

async function appendContactToExcel(fileId, contact, token) {
  const data = await graphRequest("GET", "/me/drive/items/" + fileId + "/workbook/worksheets('Sheet1')/usedRange", token);
  const nextRow = data.rowCount + 1;
  const today = new Date().toISOString().split("T")[0];
  const body = {
    values: [[contact.firstName, contact.lastName, contact.email, contact.mobile, contact.phone, contact.title, contact.company, contact.street, contact.city, contact.postal, contact.country, today, today]]
  };
  return graphRequest("PATCH", "/me/drive/items/" + fileId + "/workbook/worksheets('Sheet1')/range(address='A" + nextRow + ":M" + nextRow + "')", token, body);
}

async function findExcelRow(fileId, email, token) {
  const data = await graphRequest("GET", "/me/drive/items/" + fileId + "/workbook/worksheets('Sheet1')/usedRange", token);
  if (data && data.values) {
    for (let i = 0; i < data.values.length; i++) {
      if (data.values[i][2] === email) {
        return i + 1; // 1-based row number
      }
    }
  }
  return null;
}

async function updateContactInExcel(fileId, rowIndex, contact, token) {
  const today = new Date().toISOString().split("T")[0];
  const currentData = await graphRequest("GET", "/me/drive/items/" + fileId + "/workbook/worksheets('Sheet1')/range(address='A" + rowIndex + ":M" + rowIndex + "')", token);
  const dateAdded = currentData && currentData.values && currentData.values[0] ? currentData.values[0][11] : today;
  const body = {
    values: [[contact.firstName, contact.lastName, contact.email, contact.mobile, contact.phone, contact.title, contact.company, contact.street, contact.city, contact.postal, contact.country, dateAdded, today]]
  };
  return graphRequest("PATCH", "/me/drive/items/" + fileId + "/workbook/worksheets('Sheet1')/range(address='A" + rowIndex + ":M" + rowIndex + "')", token, body);
}

async function getAllContactsForExport(token) {
  const data = await graphRequest("GET", "/me/contacts?$top=999&$select=givenName,surname,emailAddresses,mobilePhone,businessPhones,jobTitle,companyName,businessAddress", token);
  return data && data.value ? data.value : [];
}

window.Graph = {
  searchContactByEmail: searchContactByEmail,
  createContact: createContact,
  updateContact: updateContact,
  findOrCreateExcelFile: findOrCreateExcelFile,
  ensureHeaders: ensureHeaders,
  appendContactToExcel: appendContactToExcel,
  findExcelRow: findExcelRow,
  updateContactInExcel: updateContactInExcel,
  getAllContactsForExport: getAllContactsForExport
};