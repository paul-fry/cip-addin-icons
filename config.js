// ============================================================
// CIP_CONFIG - EDIT THIS BLOCK BEFORE DEPLOYMENT
// ============================================================
// See the Setup Guide PDF (DOC-IT-006) for details.

window.CIP_CONFIG = {
  // ----- Azure AD app registration (Setup Guide section 1) -----
  clientId:   "56a3b894-c4be-44fd-8544-56930071a0ee",
  tenantId:   "7b0510a2-88be-4fdb-91fa-e3e3d7630786",

  // ----- Microsoft Graph scopes -----
  // For read-mode filing alone: User.Read, Sites.ReadWrite.All, Files.ReadWrite.All
  // Mail.Read added so the add-in can fetch the sent copy from Sent Items after Send completes
  scopes: [
    "User.Read",
    "Sites.ReadWrite.All",
    "Files.ReadWrite.All",
    "Mail.Read"
  ],

  // ----- Filtering & defaults -----
  siteFilter: [],   // empty = show all sites the tech has access to
  defaultFolderTemplate: "{YYYY}/{MM}",   // placeholders: {YYYY},{MM},{DD},{client},{project}
  preferredLibrary: "Filed Emails",

  // ----- Client domain → SharePoint destination map -----
  // When composing a new email, the add-in reads the To: line and auto-suggests
  // the matching destination using this map. The first matching entry wins.
  // Patterns can be exact ("foo.com"), wildcard ("*.gov.uk"), or substring.
  // siteHint can be a site name, hostname, or URL fragment to match against
  // the Graph site list. Add as many entries as you need.
  clientMap: [
    // Example:
    // { match: "cardiff-architects.example.com", siteHint: "cip-clients/Cardiff", libraryHint: "Filed Emails", client: "Cardiff Architects Ltd" },
    // { match: "*.gov.uk",                        siteHint: "cip-public-sector",   libraryHint: "Filed Emails", client: "Public Sector" }
  ]
};
