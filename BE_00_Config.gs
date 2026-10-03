// SiM-Guru — backend configuration
const SS_ID = SpreadsheetApp.getActiveSpreadsheet().getId();
const EXT_USERS_SS_ID = "1cnMeA7xinjvilVloS7A0_SFxzOBMON7DIiT7_cnnYUg";
const EXT_USERS_SHEET = "Users";
const SHEET_NAME = {
  USERS: "Users",
  SESSIONS: "Sessions",
  SCHEDULES: "Schedules",
  LOGS: "Teaching_Logs",
  CONFIG: "Config",
  CALENDAR: "Academic_Calendar",
  HONOR_HISTORY: "Honor_History",
  ALLOWANCES: "Allowances",
  SUBJECTS: "Subjects",
  RESET_REQUESTS: "Reset_Requests",
  ATTENDANCE: "Daily_Attendance",
  PICKET_SCHEDULES: "Picket_Schedules",
  SUBSTITUTES: "Substitutes",
  CEREMONY_SCHEDULES: "Ceremony_Schedules",
  ANNOUNCEMENTS: "Announcements",
  ATTENDANCE_SCHED_TEMPLATES: "Attendance_Sched_Templates",
  ATTENDANCE_LEAVES: "Attendance_Leaves",
  STUDENT_ATTENDANCE: "Student_Attendance",
};
const EXAM_SHEET = {
  PERIODS: "Exam_Periods",
  SESSIONS: "Exam_Sessions",
  ROOMS: "Exam_Rooms",
  SUPERVISORS: "Exam_Supervisors",
  COMMITTEE: "Exam_Committee",
  BAP: "Exam_BAP",
};
const EVENT_SHEET = {
  DEFINITIONS: "Event_Definitions",
  ATTENDANCE: "Event_Attendance",
  JOURNALS: "Event_Journals",
};
