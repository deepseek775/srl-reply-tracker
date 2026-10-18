// ================= SRL Reply Tracker - settings =================
// 1) Paste the "Application (client) ID" from your Entra app registration here:
window.TRACKER_CONFIG = {
  clientId: "PASTE-CLIENT-ID-HERE",

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
