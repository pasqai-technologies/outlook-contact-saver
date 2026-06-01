// --- BLOCK: Signature Extraction ---
// Purpose: Extract contact fields from email body text

function cleanSignature(body) {
  const text = body.replace(/<[^>]+>/g, " ").replace(/\s+/g, " ");
  const lines = text.split("\n");

  let sigStartIndex = -1;

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i].trim();

    // Exact "--" line
    if (line === "--") {
      sigStartIndex = i + 1;
      break;
    }

    // Greeting/closing patterns
    const greetingPattern = /^(regards|best regards|med vänliga hälsningar|mvh|hälsningar|cordially|kind regards|thanks|thank you)[,.]?\s*$/i;
    if (greetingPattern.test(line)) {
      sigStartIndex = i + 1;
      break;
    }

    // Name pattern (capitalized first and last name) after line 3
    if (i > 3) {
      const namePattern = /^[A-ZÀÁÂÄÆÃÅÇ][a-zàáâäæãåç]+ [A-ZÀÁÂÄÆÃÅÇ][a-zàáâäæãåç\-]+$/;
      if (namePattern.test(line)) {
        sigStartIndex = i;
        break;
      }
    }
  }

  if (sigStartIndex === -1) {
    // Return last 20 lines
    const start = Math.max(0, lines.length - 20);
    return lines.slice(start).join("\n");
  }

  return lines.slice(sigStartIndex).join("\n");
}

function extractEmail(text) {
  const match = text.match(/[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}/);
  return match ? match[0] : "";
}

function extractPhones(text) {
  const mobilePatterns = [
    /\+46\s?7[\d\s\-]{8,12}/,
    /\+44\s?7[\d\s\-]{9,11}/,
    /\b07\d{2}[\s\-]?\d{3}[\s\-]?\d{3}\b/,
    /\+\d{1,3}\s?7[\d\s\-]{7,12}/,
  ];

  let mobile = "";
  let other = "";

  // Find mobile
  for (const pattern of mobilePatterns) {
    const match = text.match(pattern);
    if (match) {
      mobile = match[0];
      break;
    }
  }

  // Find other phone (not mobile)
  const otherPattern = /\+?[\d\s\-\.\(\)]{7,20}/;
  const allMatches = [];
  let match;
  const regex = new RegExp(otherPattern.source, otherPattern.flags);
  while ((match = regex.exec(text)) !== null) {
    allMatches.push({ match: match[0], index: match.index });
  }

  for (const m of allMatches) {
    // Check if this match is a mobile pattern
    let isMobile = false;
    for (const pattern of mobilePatterns) {
      if (pattern.test(m.match)) {
        isMobile = true;
        break;
      }
    }
    if (!isMobile) {
      other = m.match;
      break;
    }
  }

  return { mobile, other };
}

function extractTitle(lines) {
  const keywords = [
    "consultant", "manager", "director", "developer", "engineer",
    "analyst", "architect", "designer", "partner", "associate",
    "specialist", "coordinator", "officer", "executive", "lead",
    "head", "vd", "ceo", "cto", "cfo", "coo", "president",
    "vice", "senior", "junior", "account",
  ];

  for (const line of lines) {
    if (line.length >= 60) continue;
    const lowerLine = line.toLowerCase();
    for (const keyword of keywords) {
      if (lowerLine.includes(keyword)) {
        return line;
      }
    }
  }

  return "";
}

function extractCompany(lines) {
  for (const line of lines) {
    const trimmed = line.trim();
    if (trimmed.length < 2 || trimmed.length > 60) continue;
    if (trimmed.includes("@")) continue;
    if (/\d/.test(trimmed) && /\d/.test(trimmed.replace(/\s/g, ""))) {
      // Check if it's mostly digits (phone-like)
      const digitCount = (trimmed.match(/\d/g) || []).length;
      if (digitCount >= 3) continue;
    }
    if (/http/.test(trimmed) || /www/.test(trimmed)) continue;
    if (/^\d/.test(trimmed)) continue; // Starts with digit (address pattern)
    if (trimmed === trimmed.toLowerCase()) continue; // All lowercase

    return trimmed;
  }

  return "";
}

function extractAddress(text) {
  const lines = text.split("\n");

  // Street
  const streetPattern = /\d+\s+[A-Za-zÀ-ÿ\s]+(street|st|road|rd|avenue|ave|lane|ln|way|drive|dr|gatan|vägen|gränd|torget|allén)\b/i;
  const streetMatch = text.match(streetPattern);
  const street = streetMatch ? streetMatch[0] : "";

  // Postal
  const postalPattern = /\b([A-Z]{1,2}\d[\dA-Z]?\s?\d[A-Z]{2}|\d{3}\s?\d{2}|\d{4,6})\b/;
  const postalMatch = text.match(postalPattern);
  const postal = postalMatch ? postalMatch[0] : "";

  // Country
  const countries = [
    "Sweden", "Sverige", "United Kingdom", "UK", "Germany",
    "Deutschland", "Norway", "Norge", "Denmark", "Danmark",
    "Finland", "Netherlands", "France", "USA", "United States",
  ];
  let country = "";
  for (const line of lines) {
    for (const c of countries) {
      if (line.toLowerCase().includes(c.toLowerCase())) {
        country = c;
        break;
      }
    }
    if (country) break;
  }

  // City: line between street and postal
  let city = "";
  const streetIndex = lines.findIndex((l) => l.includes(street));
  const postalIndex = lines.findIndex((l) => l.includes(postal));

  if (streetIndex !== -1 && postalIndex !== -1 && postalIndex > streetIndex) {
    for (let i = streetIndex + 1; i < postalIndex; i++) {
      const line = lines[i].trim();
      if (line.length >= 2 && line.length <= 40 && !/^\d/.test(line)) {
        city = line;
        break;
      }
    }
  }

  return { street, city, postal, country };
}

function extractName(senderName, signatureLines) {
  if (senderName && !senderName.includes("@")) {
    const parts = senderName.split(" ");
    if (parts.length >= 2) {
      return {
        first: parts[0],
        last: parts.slice(1).join(" "),
      };
    }
  }

  // Search signature lines for name pattern
  const namePattern =
    /^[A-ZÀÁÂÄÆÃÅÇ][a-zàáâäæãåç]+ [A-ZÀÁÂÄÆÃÅÇ][a-zàáâäæãåç\-]+$/;
  for (const line of signatureLines) {
    if (namePattern.test(line.trim())) {
      const parts = line.trim().split(" ");
      return {
        first: parts[0],
        last: parts.slice(1).join(" "),
      };
    }
  }

  return { first: senderName || "", last: "" };
}

function extractFromEmail(body, senderEmail, senderName) {
  const sigText = cleanSignature(body);
  const lines = sigText.split("\n").filter((l) => l.trim());

  const name = extractName(senderName, lines);
  const email = extractEmail(sigText) || senderEmail;
  const phones = extractPhones(sigText);
  const title = extractTitle(lines);
  const company = extractCompany(lines);
  const address = extractAddress(sigText);

  return {
    firstName: name.first,
    lastName: name.last,
    email,
    mobile: phones.mobile,
    phone: phones.other,
    title,
    company,
    street: address.street,
    city: address.city,
    postal: address.postal,
    country: address.country,
  };
}

window.Extractor = { extractFromEmail, cleanSignature };