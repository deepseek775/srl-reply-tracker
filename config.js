// ================= SRL Reply Tracker - settings =================
// 1) Paste the "Application (client) ID" from your Entra app registration here:
window.TRACKER_CONFIG = {
  clientId: "de23bf32-0c61-4fe7-b246-3ce2d03f4364",
  tenantId: "35f7a35d-a292-4c08-8d71-204da773a21a",   // SRL (Square Root Maritime Services Co.)

  // How far back to look
  daysBack: 60,          // emails YOU sent with an @mention
  inDaysBack: 30,        // emails where others @mentioned YOU

  // Highlight in red when waiting this many days or more
  redAfter: 3,

  // Weekend days (0=Sun ... 6=Sat). KSA weekend = Friday + Saturday
  weekend: [5, 6],

  // Also track EVERY email you send to these people (To line), even without @mention
  // e.g. alsoTrack: ["saad.iqbal@srlksa.com"]
  alsoTrack: [],

  // Ignore these addresses completely
  ignore: []
};
