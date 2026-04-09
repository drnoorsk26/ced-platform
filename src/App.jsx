import { useState, useEffect, useCallback, useMemo } from "react";
import { BarChart, Bar, XAxis, YAxis, Tooltip, ResponsiveContainer, PieChart, Pie, Cell, LineChart, Line, CartesianGrid, Legend } from "recharts";
import * as XLSX from "xlsx";

// ─── LOGO ─────────────────────────────────────────────────────────────
const LOGO = "data:image/jpeg;base64,/9j/4AAQSkZJRgABAQAAAQABAAD/2wBDAAUDBAQEAwUEBAQFBQUGBwwIBwcHBw8LCwkMEQ8SEhEPERETFhwXExQaFRERGCEYGh0dHx8fExciJCIeJBweHx7/2wBDAQUFBQcGBw4ICA4eFBEUHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh7/wAARCABkAGQDASIAAhEBAxEB/8QAHwAAAQUBAQEBAQEAAAAAAAAAAAECAwQFBgcICQoL/8QAtRAAAgEDAwIEAwUFBAQAAAF9AQIDAAQRBRIhMUEGE1FhByJxFDKBkaEII0KxwRVS0fAkM2JyggkKFhcYGRolJicoKSo0NTY3ODk6Q0RFRkdISUpTVFVWV1hZWmNkZWZnaGlqc3R1dnd4eXqDhIWGh4iJipKTlJWWl5iZmqKjpKWmp6ipqrKztLW2t7i5usLDxMXGx8jJytLT1NXW19jZ2uHi4+Tl5ufo6erx8vP09fb3+Pn6/8QAHwEAAwEBAQEBAQEBAQAAAAAAAAECAwQFBgcICQoL/8QAtREAAgECBAQDBAcFBAQAAQJ3AAECAxEEBSExBhJBUQdhcRMiMoEIFEKRobHBCSMzUvAVYnLRChYkNOEl8RcYGRomJygpKjU2Nzg5OkNERUZHSElKU1RVVldYWVpjZGVmZ2hpanN0dXZ3eHl6goOEhYaHiImKkpOUlZaXmJmaoqOkpaanqKmqsrO0tba3uLm6wsPExcbHyMnK0tPU1dbX2Nna4uPk5ebn6Onq8vP09fb3+Pn6/9oADAMBAAIRAxEAPwD4yooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiit/wRoTa3qqq0Es8MbKDFHnfM7HCRLjnLHPToAx7UegEeg+G73VI1uWdLS0ZtqzSAkyHuqKMs5+gwO5FdNc+BbOxshc3ya2sRKr5zQxxKSRuHylieRz9Oa9I086ZoVrFcRxQ32thQ65JjgtIlypRhkeSI5FGMct7hjnJjik1PSr65GrQ/ZrIRtMIbIbV2oyJt3sGbCuRnHoewNdcMNBfG235X/TX79+xk6j+yec6x4JurdmXT5ZZ5VXcbSeAw3GMZ+VSSH47Kc+1cmQQcEYIr3LVtTutbtLS2v720+wJdb5L2CNhsd2GTKp+ZOrkHBBY8H05T4r+GDa3l3dwt501u5EsqqdtwmR8+cANIm5VkK8EkH1NZVaPIuaLuvy/r70VGfNo9zzeiiisSwooooAKKKKACiiigAr2D4cw2Gj6MLi9jaaRbYzLbhFYTyygAK25gQuwouQCcucFW2mvIoIzLMka9XYKPx4r3eSOWw8PQzpe3ywzXU0ptDt+zt5KtsdcE8jYgIIU8DqMEbYde/zdl/X6kVNrdzm9bupXne1MzzESb7iVnLNPP8AxMWPLY+6M+hPUmuv8GJHLpdvBHFbwQI3mX++T96cfeODyFdMYxwCDzXBW0F1czw2dkk811cNtRIFLSP22qBzlj6c4Br33wj+zpPf6JpEmqeI/sST/6Tq FpaxxDfEGXIALcFxwDu4HOOnO2McIwUJSt593bt/WpNFSbulc8Gt7s2941xbxqEYsDExJVoyfuN6jHH69a7O2lh1HwvNp0kK3BsYxNaSOSZHtXLZh7lj8zptUAdWYnaoHR/H34TaT8O9L0q/0rWb2/F7cSQyLcKny7VDAgp+Oc+1cn8NJ2/teyhMjokrT2khWJ5cxvHvwY1ILjKN8vIOeQ3Q9HPTrR546rZ/1/W5laUHZ7nimsWZ0/VLmyJ3eTIUDf3hng/iMH8aqV0vxJtRa+JnUBwTGoO9drfLmPkdj8nSuary7W0Z1BRRRQAUUUUAFFFFAE+nuI7+3duiyqT+Yr2XUtShurSLS4zdNLawXSv5nl+WCdxwmFD9sncxHPAFeJivTra8X7baant3JOiTMo5yGGHH57xXRhfikutvy/4czq7JnoX7MSq3xl0rcoYCG4IyM/8ALI19fWdn5UFwqafDI8jsBGSihkzwc7ccDkAgmvhHw1b6pb+MrSw0nUpLC8knEMV3G5Xajfx5HYrz716rrWq+N9P1rRtJ8OfEfU77TtdbyRc3MaB42Qjf0ByMHIIPPTNcWZxi8TGXOldX1vsrvon0T+46MLf2TXK9/LrZdzqv2wonh8E+Go5LZLdl1CXKqynd+6+8doAyfpXgfhKe0tJoLm+uDbwrd537XYbhE2Adh3hcsMlfmAzjnFdF8ZhfxXVkG8Yaj4j02XeYmvMK0UqYDgoOBwwx0ODzXF6ijQwWtiFPmIpkdcfxvjC/UKFH1zXo5fZ4eNne7b0/4KXXyObEXVR3Wxy/xNlim1+OSFmaNoNyls5IMjsM556Edea5atjxlMsviK5RG3JBtt1P+4oU/qCax642+Zt9zW1tAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKAP/Z";

// ─── TRANSLATIONS ─────────────────────────────────────────────────────
const T = {
  en: {
    dir: "ltr", font: "'Noto Sans',sans-serif",
    appDefaultName: "Continual Education Department",
    appDefaultSub: "University of Mashreq — Baghdad, Iraq",
    login: "Sign In", loginTitle: "Sign In to Platform", username: "Username", password: "Password",
    invalidCred: "Invalid credentials. Check username and password.",
    dashboard: "Dashboard", yearlyPlans: "Yearly plans", allActivities: "All activities",
    initiatives: "Initiatives", trainers: "Trainer profiles", reports: "Event reports",
    correspondence: "Correspondence", backup: "Backup & export", access: "Access control",
    settings: "Settings", notifications: "Notifications", analytics: "Analytics",
    logout: "Logout", search: "Search...", add: "Add", save: "Save", cancel: "Cancel",
    edit: "Edit", delete: "Delete", approve: "Approve", complete: "Complete", back: "Back",
    newPlan: "New plan", newActivity: "New activity", newInitiative: "New initiative",
    newTrainer: "New trainer", newReport: "New report", newCorr: "New correspondence",
    college: "College", year: "Year", status: "Status", type: "Type", date: "Date",
    title: "Title", presenter: "Presenter", delivery: "Delivery", target: "Target audience",
    venue: "Venue", duration: "Duration", objectives: "Objectives", notes: "Notes",
    source: "Source", attendees: "Attendees", name: "Full name", phone: "Phone", email: "Email",
    workplace: "Workplace", degree: "Degree", generalSpec: "General specialization",
    specificSpec: "Specific specialization", internal: "Internal", external: "External",
    incomingCorr: "Incoming", outgoingCorr: "Outgoing", refNumber: "Reference number",
    from: "From", to: "To", subject: "Subject", all: "All",
    totalActivities: "Total activities", completed: "Completed", upcoming: "Upcoming",
    totalAttendees: "Total attendees", plansSubmitted: "Plans submitted",
    completionRate: "Completion rate", byType: "By type", bySource: "By source",
    byMonth: "Monthly distribution", byStatus: "By status",
    upcomingActivities: "Upcoming activities", noUpcoming: "No upcoming activities",
    activitiesTimeline: "Activities timeline", sourceDistribution: "Source distribution",
    statusOverview: "Status overview", monthlyTrend: "Monthly trend",
    exportBackup: "Export backup", importRestore: "Import / restore",
    exportAll: "Export all data", downloadBackup: "Download backup file",
    importData: "Import from backup", storageInfo: "Where is data stored?",
    storageDesc: "All data persists in artifact storage across sessions.",
    modifyInfo: "How to modify without losing data",
    modifyDesc: "1. Export backup first  2. Request changes  3. Import backup into new version",
    platformName: "Platform name", platformSub: "Platform subtitle", language: "Language",
    manageUsers: "Manage users", addUser: "Add user", role: "Role",
    admin: "Admin", viewer: "Viewer", collegeRep: "College rep",
    adminPerms: "Full system access, manage all data, approve plans, add users",
    viewerPerms: "Read-only access to all data, no editing",
    collegePerms: "Manage own college plan and activities only",
    mediaNotified: "Media notified", mediaLinks: "Media/published links",
    confirmed: "Confirmed", summary: "Summary / outcome",
    participantNames: "Participant names", totalParticipants: "Total participants",
    selectActivity: "Select completed activity", noReports: "No reports yet",
    noPlans: "No plans found", noActivities: "No activities match filters",
    noCorr: "No correspondence records", noTrainers: "No trainer profiles",
    export: "Export", print: "Print", dob: "Date of birth", gender: "Gender",
    male: "Male", female: "Female", inPerson: "In-Person", online: "Online", hybrid: "Hybrid",
    students: "Students", faculty: "Faculty", staff: "Staff", externalAud: "External Audience",
    incoming: "Incoming (واردة)", outgoing: "Outgoing (صادرة)",
    received: "Received", processed: "Processed", sent: "Sent", pending: "Pending", archived: "Archived",
    planned: "Planned", approved: "Approved", inProgress: "In Progress",
    completedStat: "Completed", postponed: "Postponed", cancelled: "Cancelled",
    pendingReview: "Pending Review",
    upcomingNotif: "activities in the next 7 days", noNotif: "No upcoming activities this week",
    guideline10days: "Per guidelines: submit at least 10 days before event date",
    submittedDate: "Submitted", acts: "activities",
    collegePlan: "College Plan", collegeRequest: "College Request",
    deptInit: "Department Initiative", univInit: "University Initiative", minDir: "Ministry Directive",
    lecture: "Lecture", workshop: "Workshop", training: "Training",
    awareness: "Awareness Lecture", seminar: "Seminar", conference: "Conference", jointProg: "Joint Program",
    results: "results", filterBy: "Filter", exportCSV: "Export CSV",
    academicYear: "Academic year 2025–2026",
    trainerDB: "Credentials database (استمارة ترشيح)",
    corrSubtitle: "الصادرة والواردة — Incoming & outgoing documents",
    planSubtitle: "College-submitted annual plans (المنهاج التدريبي السنوي)",
    initSubtitle: "Department, university, ministry, or ad-hoc college requests",
    reportSubtitle: "Post-event documentation and confirmation",
    settingsSubtitle: "Customize platform name, language, and manage users",
    analyticsSubtitle: "Detailed charts and statistics",
    notifSubtitle: "Activities happening soon",
    backupSubtitle: "Export, import, and manage your data",
    accessSubtitle: "User roles and permissions",
    mediaCoverage: "Media coverage", mediaStatus: "Media status", mediaContactedDate: "Media contacted date",
    mediaPublishedDate: "Published date", mediaChannel: "Channel/platform", mediaCoverageNotes: "Coverage notes",
    notContacted: "Not contacted", contacted: "Contacted", coverageScheduled: "Coverage scheduled",
    published: "Published", declined: "Declined",
    stickyNotes: "Quick notes", addNote: "Add note", noteText: "Type a reminder or note...",
    noNotes: "No notes yet", noteAdded: "Note added", deleteNote: "Delete note",
    platformLogo: "Platform logo", uploadLogo: "Upload logo image",
    logoHelp: "Upload a square image (PNG/JPG). It will be used on the sidebar and login page.",
    resetLogo: "Reset to default",
    uploadFile: "Upload file / image", extracting: "Extracting information...",
    extractFromFile: "Extract from document", orFillManually: "Or fill the form manually below",
    extractedFields: "Extracted fields", customField: "Custom field", fieldName: "Field name",
    fieldValue: "Value", addCustomField: "Add field", removeField: "Remove",
    uploadHint: "Upload a document, image, or photo to auto-extract activity details",
    extractionDone: "Extraction complete — review and edit below",
    extractionError: "Could not extract. Please fill manually.",
    camera: "Camera", fileUpload: "File",
    exportReport: "Export report", exportOptions: "Export options",
    exportAsWord: "Word (.docx)", exportAsExcel: "Excel (.xlsx)", exportAsPDF: "Print / PDF",
    exportBulkExcel: "Export all to Excel", includeOptions: "Include in report:",
    incHeader: "University header & logo", incDetails: "Activity full details",
    incCustomFields: "Custom/extracted fields", incMedia: "Media coverage status",
    incParticipants: "Participant names", incObjectives: "Objectives & notes",
    incAttendees: "Attendee count", incTimeline: "Date & timeline info",
    selectAll: "Select all", deselectAll: "Deselect all",
    generating: "Generating...", reportGenerated: "Report generated",
    activityReport: "Activity Report", preparedBy: "Prepared by",
    singleReport: "Single activity", bulkReport: "All activities list",
    bulkImport: "Bulk import from file", bulkImportHint: "Upload a yearly plan document to extract all activities at once",
    extractingBulk: "Extracting activities...", activitiesFound: "activities found",
    importAll: "Import all", department: "Department", noDept: "No department",
    filterExport: "Filter before export", dateFrom: "From date", dateTo: "To date",
    exportFiltered: "Export filtered", chooseFormat: "Choose format",
    auditLog: "Audit log", auditAction: "Action", auditUser: "User", auditTime: "Time",
    auditDetails: "Details", noAuditEntries: "No audit entries yet",
    loginSuccess: "Login successful", loginFailed: "Failed login attempt",
    activityCreated: "Activity created", activityUpdated: "Activity updated",
    activityDeleted: "Activity deleted", planCreated: "Plan created", planApproved: "Plan approved",
    userCreated: "User created", userDeleted: "User deleted", dataExported: "Data exported",
    dataImported: "Data imported", sessionTimeout: "Session expired",
    securityNotice: "Security notice",
    securityMsg: "This platform runs client-side in your browser. For high-sensitivity data (ministry correspondence, confidential documents), contact your IT department about hosting this as a server-side application. Use this platform for workflow coordination and scheduling. Avoid storing sensitive personal information or classified documents.",
    dismissNotice: "I understand", changePassword: "Change password",
    currentPassword: "Current password", newPassword: "New password",
    confirmPassword: "Confirm new password", passwordChanged: "Password changed successfully",
    passwordMismatch: "Passwords do not match", wrongPassword: "Current password is incorrect",
    sessionExpired: "Your session has expired. Please sign in again.",
    lastLogin: "Last login", activeSession: "Active session",
    forgotPassword: "Forgot password?", recoveryKey: "Recovery key",
    recoveryKeyDesc: "This is your recovery key. Write it down and keep it in a safe place. You will need it to reset your password if you ever forget it. This key is shown ONLY ONCE.",
    iSavedIt: "I saved it securely", enterRecoveryKey: "Enter your recovery key",
    resetPassword: "Reset password", recoverySuccess: "Password reset! You can now sign in.",
    recoveryFailed: "Invalid recovery key.", noRecoveryKey: "No recovery key has been set up.",
    setupRecovery: "Set up recovery key", generateKey: "Generate recovery key",
    recoveryKeyExists: "Recovery key is set up", regenerateKey: "Regenerate key",
    recoveryWarning: "Keep this key safe — it is the ONLY way to recover your account.",
    newPasswordAfterRecovery: "New password", selectUser: "Select your username",
    attachedFile: "Attached file", downloadFile: "Download file",
    corrExtractHint: "Upload a letter, memo, or official document to auto-extract details",
    bulkCorrImport: "Bulk import correspondence", bulkCorrHint: "Upload a document with multiple letters/memos to extract all at once",
    trainerExtractHint: "Upload a CV, nomination form, or credentials document to auto-extract",
    bulkTrainerImport: "Bulk import trainers", bulkTrainerHint: "Upload a list of trainers/lecturers to extract all profiles at once",
    trainersFound: "trainer profiles found",
    storageManager: "Storage manager", storageUsage: "Storage usage",
    storageKey: "Data category", storageSize: "Size", storageLimit: "Limit: 5MB per key",
    storageTotal: "Total used", storageWarning: "Warning: approaching limit",
    storageDanger: "Danger: near capacity — export backup and remove old file attachments",
    storageSafe: "Storage healthy", clearAttachments: "Clear file attachments",
    clearAttachmentsDesc: "Remove stored files from this category to free space. Data fields are kept.",
    storageFreed: "Attachments cleared",
    recordCount: "records", withFiles: "with files",
    offloadFiles: "Export files to device", offloadDesc: "Download all attached files as separate downloads, then clear them from platform storage to free space",
    offloadAndClear: "Download all & clear from storage", offloadOnly: "Download only (keep in storage)",
    downloading: "Downloading files...", filesDownloaded: "files downloaded",
    externalRef: "External reference", externalRefHint: "Store a reference to where the file is saved (e.g. 'USB-2026/Folder3/letter.pdf' or 'Google Drive > CED > Corr')",
    noFilesStored: "No file attachments stored",
    firstTimeSetup: "Welcome — First time setup",
    setupDesc: "Create your administrator account. These credentials are not stored in the source code — only you will know them.",
    setupAdminName: "Your full name", setupUsername: "Choose a username",
    setupPassword: "Choose a password (min 6 characters)", setupConfirm: "Confirm password",
    setupComplete: "Create account & start", setupError: "Please fill all fields correctly",
    setupPasswordShort: "Password must be at least 6 characters",
    setupPasswordMatch: "Passwords do not match",
    noCredentialsInCode: "No credentials are stored in the source code",
    manageColleges: "Manage colleges & departments", addCollege: "Add college/department",
    collegeName: "College or department name", removeCollege: "Remove",
    resetColleges: "Reset to default list", collegesUpdated: "List updated",
    collegeCount: "colleges/departments",
    archiveManager: "Archive & extend storage", archiveDesc: "Move completed academic years to archive storage. Each archived year gets its own 5MB key.",
    archiveYear: "Archive year", activeData: "Active data", archivedYears: "Archived years",
    loadArchive: "Load", archiveLoaded: "Archive loaded into view",
    archiveSaved: "Year archived — storage freed", deleteArchive: "Delete archive",
    selectiveExport: "Selective export", exportSelected: "Export selected",
    includeInExport: "Include", plansData: "Yearly plans", initData: "Initiatives",
    trainersData: "Trainers", corrData: "Correspondence", reportsData: "Reports",
    statsSnapshot: "Stats preserved", archiveNote: "Archived data can be loaded for viewing or exported anytime",
    attendanceSource: "Attendance synced from event reports", reportSyncsActivity: "Filing a report auto-updates the activity: attendees, participants, and marks it Completed",
    overdue: "Overdue", overdueDesc: "Past the scheduled date but not completed",
    clickToView: "Click metric to view details", hideDetails: "Hide details",
    directivesPage: "Directives & Orders", directivesSubtitle: "Cross-college orders, coordination, and compliance tracking",
    newDirective: "New directive", directiveTitle: "Directive title", directiveDesc: "Description / instructions",
    directiveType: "Type", deadline: "Deadline", targetColleges: "Target colleges", allColleges: "All colleges",
    selectColleges: "Select colleges", compliance: "Compliance", compliant: "Compliant", notCompliant: "Not started",
    inProgress: "In progress", linkActivity: "Link activity", linkedActivity: "Linked activity",
    createForColleges: "Create initiative for each college", directiveSaved: "Directive saved",
    overallCompliance: "Overall compliance", collegeCompliance: "College compliance",
    markCompliant: "Mark compliant", directiveSource: "Directive source",
    staffPage: "College Staff", staffSubtitle: "Staff activity tracking, appreciation letters, and performance",
    staffMember: "Staff member", appreciationLetter: "Letter of appreciation", letterIssued: "كتاب شكر issued",
    letterDate: "Letter date", letterRef: "Letter reference", addLetter: "Add appreciation letter",
    activitiesAsPresenter: "Activities as presenter", noStaffData: "No staff activity found",
    extractFromActivities: "Extract presenters from activities", extracted: "extracted",
    viewActivity: "View activity details", staffOf: "Staff of",
  },
  ar: {
    dir: "rtl", font: "'Noto Sans Arabic','Noto Sans',sans-serif",
    appDefaultName: "شعبة التعليم المستمر",
    appDefaultSub: "جامعة المشرق — بغداد، العراق",
    login: "تسجيل الدخول", loginTitle: "تسجيل الدخول إلى المنصة", username: "اسم المستخدم", password: "كلمة المرور",
    invalidCred: "بيانات الاعتماد غير صحيحة.",
    dashboard: "لوحة التحكم", yearlyPlans: "الخطط السنوية", allActivities: "جميع الأنشطة",
    initiatives: "المبادرات", trainers: "ملفات المدربين", reports: "تقارير الأنشطة",
    correspondence: "الصادرة والواردة", backup: "النسخ الاحتياطي", access: "إدارة الوصول",
    settings: "الإعدادات", notifications: "الإشعارات", analytics: "التحليلات",
    logout: "تسجيل الخروج", search: "بحث...", add: "إضافة", save: "حفظ", cancel: "إلغاء",
    edit: "تعديل", delete: "حذف", approve: "موافقة", complete: "إكمال", back: "رجوع",
    newPlan: "خطة جديدة", newActivity: "نشاط جديد", newInitiative: "مبادرة جديدة",
    newTrainer: "مدرب جديد", newReport: "تقرير جديد", newCorr: "مراسلة جديدة",
    college: "الكلية", year: "السنة", status: "الحالة", type: "النوع", date: "التاريخ",
    title: "العنوان", presenter: "المحاضر", delivery: "آلية التنفيذ", target: "الفئة المستهدفة",
    venue: "المكان", duration: "المدة", objectives: "الأهداف", notes: "ملاحظات",
    source: "المصدر", attendees: "الحاضرون", name: "الاسم الكامل", phone: "الهاتف", email: "البريد الإلكتروني",
    workplace: "مكان العمل", degree: "الشهادة", generalSpec: "التخصص العام",
    specificSpec: "التخصص الدقيق", internal: "داخلي", external: "خارجي",
    incomingCorr: "واردة", outgoingCorr: "صادرة", refNumber: "رقم المرجع",
    from: "من", to: "إلى", subject: "الموضوع", all: "الكل",
    totalActivities: "إجمالي الأنشطة", completed: "مكتملة", upcoming: "قادمة",
    totalAttendees: "إجمالي الحاضرين", plansSubmitted: "خطط مقدمة",
    completionRate: "نسبة الإنجاز", byType: "حسب النوع", bySource: "حسب المصدر",
    byMonth: "التوزيع الشهري", byStatus: "حسب الحالة",
    upcomingActivities: "الأنشطة القادمة", noUpcoming: "لا توجد أنشطة قادمة",
    activitiesTimeline: "الجدول الزمني للأنشطة", sourceDistribution: "توزيع المصادر",
    statusOverview: "نظرة عامة على الحالة", monthlyTrend: "الاتجاه الشهري",
    exportBackup: "تصدير النسخة الاحتياطية", importRestore: "استيراد / استعادة",
    exportAll: "تصدير جميع البيانات", downloadBackup: "تحميل ملف النسخة الاحتياطية",
    importData: "استيراد من نسخة احتياطية", storageInfo: "أين يتم تخزين البيانات؟",
    storageDesc: "جميع البيانات محفوظة في تخزين المنصة عبر الجلسات.",
    modifyInfo: "كيفية التعديل بدون فقدان البيانات",
    modifyDesc: "١. تصدير النسخة الاحتياطية أولاً  ٢. طلب التعديلات  ٣. استيراد النسخة الاحتياطية",
    platformName: "اسم المنصة", platformSub: "العنوان الفرعي", language: "اللغة",
    manageUsers: "إدارة المستخدمين", addUser: "إضافة مستخدم", role: "الدور",
    admin: "مدير", viewer: "مشاهد", collegeRep: "ممثل كلية",
    adminPerms: "وصول كامل للنظام، إدارة جميع البيانات، الموافقة على الخطط",
    viewerPerms: "وصول للقراءة فقط لجميع البيانات",
    collegePerms: "إدارة خطة وأنشطة الكلية الخاصة فقط",
    mediaNotified: "تم إبلاغ الإعلام", mediaLinks: "روابط المنشورات الإعلامية",
    confirmed: "مؤكد", summary: "الملخص / النتائج",
    participantNames: "أسماء المشاركين", totalParticipants: "إجمالي المشاركين",
    selectActivity: "اختر نشاطاً مكتملاً", noReports: "لا توجد تقارير بعد",
    noPlans: "لا توجد خطط", noActivities: "لا توجد أنشطة مطابقة",
    noCorr: "لا توجد مراسلات", noTrainers: "لا توجد ملفات مدربين",
    export: "تصدير", print: "طباعة", dob: "تاريخ الميلاد", gender: "الجنس",
    male: "ذكر", female: "أنثى", inPerson: "حضوري", online: "إلكتروني", hybrid: "مختلط",
    students: "طلبة", faculty: "أساتذة", staff: "موظفون", externalAud: "جمهور خارجي",
    incoming: "واردة", outgoing: "صادرة",
    received: "مستلم", processed: "معالج", sent: "مرسل", pending: "قيد الانتظار", archived: "مؤرشف",
    planned: "مخطط", approved: "موافق عليه", inProgress: "قيد التنفيذ",
    completedStat: "مكتمل", postponed: "مؤجل", cancelled: "ملغي",
    pendingReview: "قيد المراجعة",
    upcomingNotif: "أنشطة خلال الأيام السبعة القادمة", noNotif: "لا توجد أنشطة هذا الأسبوع",
    guideline10days: "حسب الدليل: يجب تقديم الطلب قبل 10 أيام على الأقل",
    submittedDate: "تاريخ التقديم", acts: "أنشطة",
    collegePlan: "خطة الكلية", collegeRequest: "طلب كلية",
    deptInit: "مبادرة القسم", univInit: "مبادرة الجامعة", minDir: "توجيه وزاري",
    lecture: "محاضرة", workshop: "ورشة عمل", training: "تدريب",
    awareness: "محاضرة توعوية", seminar: "ندوة", conference: "مؤتمر", jointProg: "برنامج مشترك",
    results: "نتيجة", filterBy: "تصفية", exportCSV: "تصدير CSV",
    academicYear: "العام الأكاديمي ٢٠٢٥-٢٠٢٦",
    trainerDB: "قاعدة بيانات المؤهلات (استمارة ترشيح)",
    corrSubtitle: "الصادرة والواردة — المستندات الرسمية",
    planSubtitle: "الخطط السنوية المقدمة من الكليات (المنهاج التدريبي السنوي)",
    initSubtitle: "مبادرات القسم والجامعة والوزارة وطلبات الكليات",
    reportSubtitle: "توثيق ما بعد الحدث والتأكيد",
    settingsSubtitle: "تخصيص اسم المنصة واللغة وإدارة المستخدمين",
    analyticsSubtitle: "رسوم بيانية وإحصائيات تفصيلية",
    notifSubtitle: "الأنشطة القادمة قريباً",
    backupSubtitle: "تصدير واستيراد وإدارة البيانات",
    accessSubtitle: "أدوار المستخدمين والصلاحيات",
    mediaCoverage: "التغطية الإعلامية", mediaStatus: "حالة الإعلام", mediaContactedDate: "تاريخ التواصل مع الإعلام",
    mediaPublishedDate: "تاريخ النشر", mediaChannel: "القناة/المنصة", mediaCoverageNotes: "ملاحظات التغطية",
    notContacted: "لم يتم التواصل", contacted: "تم التواصل", coverageScheduled: "تغطية مجدولة",
    published: "منشور", declined: "مرفوض",
    stickyNotes: "ملاحظات سريعة", addNote: "إضافة ملاحظة", noteText: "اكتب تذكيراً أو ملاحظة...",
    noNotes: "لا توجد ملاحظات", noteAdded: "تمت الإضافة", deleteNote: "حذف الملاحظة",
    platformLogo: "شعار المنصة", uploadLogo: "رفع صورة الشعار",
    logoHelp: "ارفع صورة مربعة (PNG/JPG). ستُستخدم في الشريط الجانبي وصفحة تسجيل الدخول.",
    resetLogo: "إعادة تعيين للافتراضي",
    uploadFile: "رفع ملف / صورة", extracting: "جاري استخراج المعلومات...",
    extractFromFile: "استخراج من المستند", orFillManually: "أو املأ النموذج يدوياً أدناه",
    extractedFields: "الحقول المستخرجة", customField: "حقل مخصص", fieldName: "اسم الحقل",
    fieldValue: "القيمة", addCustomField: "إضافة حقل", removeField: "إزالة",
    uploadHint: "ارفع مستنداً أو صورة أو صورة ملتقطة لاستخراج تفاصيل النشاط تلقائياً",
    extractionDone: "اكتمل الاستخراج — راجع وعدّل أدناه",
    extractionError: "تعذر الاستخراج. يرجى الملء يدوياً.",
    camera: "الكاميرا", fileUpload: "ملف",
    exportReport: "تصدير التقرير", exportOptions: "خيارات التصدير",
    exportAsWord: "وورد (.docx)", exportAsExcel: "إكسل (.xlsx)", exportAsPDF: "طباعة / PDF",
    exportBulkExcel: "تصدير الكل إلى إكسل", includeOptions: "تضمين في التقرير:",
    incHeader: "ترويسة الجامعة والشعار", incDetails: "تفاصيل النشاط الكاملة",
    incCustomFields: "الحقول المخصصة/المستخرجة", incMedia: "حالة التغطية الإعلامية",
    incParticipants: "أسماء المشاركين", incObjectives: "الأهداف والملاحظات",
    incAttendees: "عدد الحاضرين", incTimeline: "التاريخ والجدول الزمني",
    selectAll: "تحديد الكل", deselectAll: "إلغاء التحديد",
    generating: "جاري الإنشاء...", reportGenerated: "تم إنشاء التقرير",
    activityReport: "تقرير النشاط", preparedBy: "إعداد",
    singleReport: "نشاط واحد", bulkReport: "قائمة جميع الأنشطة",
    bulkImport: "استيراد جماعي من ملف", bulkImportHint: "ارفع مستند الخطة السنوية لاستخراج جميع الأنشطة دفعة واحدة",
    extractingBulk: "جاري استخراج الأنشطة...", activitiesFound: "أنشطة تم العثور عليها",
    importAll: "استيراد الكل", department: "القسم", noDept: "بدون قسم",
    filterExport: "تصفية قبل التصدير", dateFrom: "من تاريخ", dateTo: "إلى تاريخ",
    exportFiltered: "تصدير المصفى", chooseFormat: "اختر التنسيق",
    auditLog: "سجل المراجعة", auditAction: "الإجراء", auditUser: "المستخدم", auditTime: "الوقت",
    auditDetails: "التفاصيل", noAuditEntries: "لا توجد إدخالات في السجل",
    loginSuccess: "تسجيل دخول ناجح", loginFailed: "محاولة دخول فاشلة",
    activityCreated: "تم إنشاء نشاط", activityUpdated: "تم تحديث نشاط",
    activityDeleted: "تم حذف نشاط", planCreated: "تم إنشاء خطة", planApproved: "تمت الموافقة على الخطة",
    userCreated: "تم إنشاء مستخدم", userDeleted: "تم حذف مستخدم", dataExported: "تم تصدير البيانات",
    dataImported: "تم استيراد البيانات", sessionTimeout: "انتهت الجلسة",
    securityNotice: "تنبيه أمني",
    securityMsg: "هذه المنصة تعمل في المتصفح. للبيانات عالية الحساسية (مراسلات الوزارة، وثائق سرية)، تواصل مع قسم تقنية المعلومات لاستضافة التطبيق على خادم مؤمن. استخدم هذه المنصة لتنسيق سير العمل والجدولة. تجنب تخزين المعلومات الشخصية الحساسة أو الوثائق السرية.",
    dismissNotice: "فهمت", changePassword: "تغيير كلمة المرور",
    currentPassword: "كلمة المرور الحالية", newPassword: "كلمة المرور الجديدة",
    confirmPassword: "تأكيد كلمة المرور", passwordChanged: "تم تغيير كلمة المرور بنجاح",
    passwordMismatch: "كلمات المرور غير متطابقة", wrongPassword: "كلمة المرور الحالية غير صحيحة",
    sessionExpired: "انتهت جلستك. يرجى تسجيل الدخول مرة أخرى.",
    lastLogin: "آخر تسجيل دخول", activeSession: "جلسة نشطة",
    forgotPassword: "نسيت كلمة المرور؟", recoveryKey: "مفتاح الاسترداد",
    recoveryKeyDesc: "هذا هو مفتاح الاسترداد الخاص بك. اكتبه واحتفظ به في مكان آمن. ستحتاجه لإعادة تعيين كلمة المرور إذا نسيتها. يُعرض هذا المفتاح مرة واحدة فقط.",
    iSavedIt: "لقد حفظته بأمان", enterRecoveryKey: "أدخل مفتاح الاسترداد",
    resetPassword: "إعادة تعيين كلمة المرور", recoverySuccess: "تم إعادة تعيين كلمة المرور! يمكنك تسجيل الدخول الآن.",
    recoveryFailed: "مفتاح الاسترداد غير صحيح.", noRecoveryKey: "لم يتم إعداد مفتاح استرداد.",
    setupRecovery: "إعداد مفتاح الاسترداد", generateKey: "إنشاء مفتاح استرداد",
    recoveryKeyExists: "مفتاح الاسترداد مُعد", regenerateKey: "إعادة إنشاء المفتاح",
    recoveryWarning: "احتفظ بهذا المفتاح — إنه الطريقة الوحيدة لاستعادة حسابك.",
    newPasswordAfterRecovery: "كلمة المرور الجديدة", selectUser: "اختر اسم المستخدم",
    attachedFile: "ملف مرفق", downloadFile: "تحميل الملف",
    corrExtractHint: "ارفع خطاباً أو مذكرة أو وثيقة رسمية لاستخراج التفاصيل تلقائياً",
    bulkCorrImport: "استيراد مراسلات جماعي", bulkCorrHint: "ارفع مستنداً يحتوي عدة خطابات/مذكرات لاستخراجها دفعة واحدة",
    trainerExtractHint: "ارفع سيرة ذاتية أو استمارة ترشيح لاستخراج البيانات تلقائياً",
    bulkTrainerImport: "استيراد مدربين جماعي", bulkTrainerHint: "ارفع قائمة مدربين/محاضرين لاستخراج جميع الملفات الشخصية",
    trainersFound: "ملفات مدربين تم العثور عليها",
    storageManager: "إدارة التخزين", storageUsage: "استخدام التخزين",
    storageKey: "فئة البيانات", storageSize: "الحجم", storageLimit: "الحد: ٥ ميغابايت لكل مفتاح",
    storageTotal: "الإجمالي المستخدم", storageWarning: "تحذير: اقتراب من الحد",
    storageDanger: "خطر: قريب من السعة — صدّر نسخة احتياطية واحذف المرفقات القديمة",
    storageSafe: "التخزين سليم", clearAttachments: "مسح المرفقات",
    clearAttachmentsDesc: "إزالة الملفات المخزنة من هذه الفئة لتحرير المساحة. حقول البيانات تبقى.",
    storageFreed: "تم مسح المرفقات",
    recordCount: "سجل", withFiles: "مع ملفات",
    offloadFiles: "تصدير الملفات إلى الجهاز", offloadDesc: "تحميل جميع الملفات المرفقة كملفات منفصلة، ثم مسحها من تخزين المنصة لتحرير المساحة",
    offloadAndClear: "تحميل الكل ومسح من التخزين", offloadOnly: "تحميل فقط (إبقاء في التخزين)",
    downloading: "جاري تحميل الملفات...", filesDownloaded: "ملفات تم تحميلها",
    externalRef: "مرجع خارجي", externalRefHint: "سجل مكان حفظ الملف (مثال: 'USB-2026/مجلد3/خطاب.pdf' أو 'Google Drive > CED')",
    noFilesStored: "لا توجد ملفات مرفقة مخزنة",
    firstTimeSetup: "مرحباً — الإعداد الأولي",
    setupDesc: "أنشئ حساب المدير الخاص بك. بيانات الاعتماد لا تُخزن في الكود المصدري — أنت وحدك من يعرفها.",
    setupAdminName: "اسمك الكامل", setupUsername: "اختر اسم مستخدم",
    setupPassword: "اختر كلمة مرور (٦ أحرف على الأقل)", setupConfirm: "تأكيد كلمة المرور",
    setupComplete: "إنشاء الحساب والبدء", setupError: "يرجى ملء جميع الحقول بشكل صحيح",
    setupPasswordShort: "كلمة المرور يجب أن تكون ٦ أحرف على الأقل",
    setupPasswordMatch: "كلمات المرور غير متطابقة",
    noCredentialsInCode: "لا توجد بيانات اعتماد مخزنة في الكود المصدري",
    manageColleges: "إدارة الكليات والأقسام", addCollege: "إضافة كلية/قسم",
    collegeName: "اسم الكلية أو القسم", removeCollege: "إزالة",
    resetColleges: "إعادة تعيين للقائمة الافتراضية", collegesUpdated: "تم تحديث القائمة",
    collegeCount: "كلية/قسم",
    archiveManager: "الأرشيف وتوسيع التخزين", archiveDesc: "نقل السنوات الأكاديمية المكتملة إلى تخزين أرشيفي. كل سنة مؤرشفة تحصل على ٥ ميغابايت خاصة.",
    archiveYear: "أرشفة السنة", activeData: "البيانات النشطة", archivedYears: "السنوات المؤرشفة",
    loadArchive: "تحميل", archiveLoaded: "تم تحميل الأرشيف للعرض",
    archiveSaved: "تم أرشفة السنة — تم تحرير التخزين", deleteArchive: "حذف الأرشيف",
    selectiveExport: "تصدير انتقائي", exportSelected: "تصدير المحدد",
    includeInExport: "تضمين", plansData: "الخطط السنوية", initData: "المبادرات",
    trainersData: "المدربين", corrData: "المراسلات", reportsData: "التقارير",
    statsSnapshot: "الإحصائيات محفوظة", archiveNote: "يمكن تحميل البيانات المؤرشفة للعرض أو التصدير في أي وقت",
    attendanceSource: "الحضور محدّث من تقارير الأحداث", reportSyncsActivity: "تقديم التقرير يحدّث النشاط تلقائياً: الحضور والمشاركين ويضعه كمكتمل",
    overdue: "متأخر", overdueDesc: "تجاوز التاريخ المحدد ولم يُكتمل",
    clickToView: "اضغط على المقياس لعرض التفاصيل", hideDetails: "إخفاء التفاصيل",
    directivesPage: "الأوامر والتوجيهات", directivesSubtitle: "أوامر عبر الكليات والتنسيق وتتبع الامتثال",
    newDirective: "توجيه جديد", directiveTitle: "عنوان التوجيه", directiveDesc: "الوصف / التعليمات",
    directiveType: "النوع", deadline: "الموعد النهائي", targetColleges: "الكليات المستهدفة", allColleges: "جميع الكليات",
    selectColleges: "اختر الكليات", compliance: "الامتثال", compliant: "مكتمل", notCompliant: "لم يبدأ",
    inProgress: "قيد التنفيذ", linkActivity: "ربط نشاط", linkedActivity: "النشاط المرتبط",
    createForColleges: "إنشاء مبادرة لكل كلية", directiveSaved: "تم حفظ التوجيه",
    overallCompliance: "الامتثال العام", collegeCompliance: "امتثال الكليات",
    markCompliant: "تحديد كمكتمل", directiveSource: "مصدر التوجيه",
    staffPage: "كادر الكليات", staffSubtitle: "تتبع أنشطة الكادر وكتب الشكر والأداء",
    staffMember: "عضو كادر", appreciationLetter: "كتاب شكر", letterIssued: "تم إصدار كتاب شكر",
    letterDate: "تاريخ الكتاب", letterRef: "رقم الكتاب", addLetter: "إضافة كتاب شكر",
    activitiesAsPresenter: "الأنشطة كمحاضر", noStaffData: "لا توجد بيانات كادر",
    extractFromActivities: "استخراج المحاضرين من الأنشطة", extracted: "تم استخراجهم",
    viewActivity: "عرض تفاصيل النشاط", staffOf: "كادر",
  }
};

// ─── CONSTANTS ────────────────────────────────────────────────────────
const DEFAULT_COLLEGES = ["College of Dentistry","College of Pharmacy","College of Law","College of Technical Engineering","College of Fine Arts","College of Medical Sciences Technologies","College of Administrative Sciences","College of Design","College of Medical Physics","College of Medical Laboratory Technologies","College of Anaesthesiology","Department of Accounting","Department of Business Administration","Department of Finance & Banking"];
const ACT_TYPES = ["Lecture","Workshop","Training","Awareness Lecture","Seminar","Conference","Joint Program"];
const SOURCES = ["College Plan","College Request","Department Initiative","University Initiative","Ministry Directive"];
const STATUSES = ["Planned","Approved","In Progress","Completed","Postponed","Cancelled"];
const DELIVERY = ["In-Person","Online","Hybrid"];
const TARGETS = ["Students","Faculty","Staff","External Audience","All"];
const MONTHS_EN = ["Sep","Oct","Nov","Dec","Jan","Feb","Mar","Apr","May","Jun","Jul","Aug"];
const MONTHS_AR = ["سبتمبر","أكتوبر","نوفمبر","ديسمبر","يناير","فبراير","مارس","أبريل","مايو","يونيو","يوليو","أغسطس"];
const MONTH_MAP = [4,5,6,7,8,9,10,11,0,1,2,3]; // maps calendar month (0=Jan) to academic index (0=Sep)
const PIE_COLORS = ["#1a5276","#c9a84c","#27ae60","#e74c3c","#8e44ad","#e67e22","#2980b9","#16a085"];

const DEFAULT_USERS = [];

const SAMPLE_PLANS = [];
const SAMPLE_INIT = [];
const SAMPLE_TRAINERS = [];
const SAMPLE_CORR = [];

// ─── STORAGE ──────────────────────────────────────────────────────────
const K = {session:"cedp-session",plans:"cedp-plans",init:"cedp-init",trainers:"cedp-trainers",corr:"cedp-corr",reports:"cedp-reports",users:"cedp-users",config:"cedp-config",notes:"cedp-notes",audit:"cedp-audit",recovery:"cedp-recovery",roster:"cedp-roster",directives:"cedp-directives",staffMerges:"cedp-staff-merges"};
// Storage functions — use localStorage for GitHub Pages compatibility
async function sG(k,fb){try{const r=localStorage.getItem(k);return r?JSON.parse(r):fb;}catch{return fb;}}
async function sS(k,v){try{localStorage.setItem(k,JSON.stringify(v));}catch(e){console.error(e);}}
// Personal storage for session only (login state per user)
async function pG(k,fb){try{const r=localStorage.getItem(k);return r?JSON.parse(r):fb;}catch{return fb;}}
async function pS(k,v){try{localStorage.setItem(k,JSON.stringify(v));}catch(e){console.error(e);}}

// ─── SECURITY UTILITIES ──────────────────────────────────────────────
async function hashPassword(password) {
  const encoder = new TextEncoder();
  const data = encoder.encode(password + "cedp-salt-uom-2026");
  const hashBuffer = await crypto.subtle.digest("SHA-256", data);
  const hashArray = Array.from(new Uint8Array(hashBuffer));
  return hashArray.map(b => b.toString(16).padStart(2, "0")).join("");
}

async function verifyPassword(password, hash) {
  const computed = await hashPassword(password);
  return computed === hash;
}

const SESSION_TIMEOUT_MS = 30 * 60 * 1000; // 30 minutes

async function addAuditEntry(action, user, details) {
  const log = await sG(K.audit, []);
  const entry = { id: gid(), timestamp: new Date().toISOString(), user: user?.name || "System", role: user?.role || "", action, details: details || "" };
  const updated = [entry, ...log].slice(0, 500);
  await sS(K.audit, updated);
  return updated;
}

function generateRecoveryKey() {
  const chars = "ABCDEFGHJKLMNPQRSTUVWXYZ23456789";
  let key = "";
  for (let i = 0; i < 20; i++) {
    if (i > 0 && i % 5 === 0) key += "-";
    key += chars[Math.floor(Math.random() * chars.length)];
  }
  return key;
}

async function saveRecoveryKeyHash(plainKey) {
  const hash = await hashPassword(plainKey.replace(/-/g, ""));
  await sS(K.recovery, { hash, createdAt: new Date().toISOString() });
}

async function verifyRecoveryKey(plainKey) {
  const stored = await sG(K.recovery, null);
  if (!stored || !stored.hash) return false;
  return await verifyPassword(plainKey.replace(/-/g, ""), stored.hash);
}

// ─── AI EXTRACTION ────────────────────────────────────────────────────
function makeApiHeaders() {
  const h = { "Content-Type": "application/json", "anthropic-version": "2023-06-01" };
  try {
    const k = typeof window !== "undefined" && localStorage.getItem("ced_api_key");
    if (k) { h["x-api-key"] = k; h["anthropic-dangerous-direct-browser-access"] = "true"; }
  } catch {}
  return h;
}

async function extractFromFile(base64Data, mediaType, lang) {
  const isImage = mediaType.startsWith("image/");
  const isPdf = mediaType === "application/pdf";
  const content = [];
  if (isImage) {
    content.push({ type: "image", source: { type: "base64", media_type: mediaType, data: base64Data } });
  } else if (isPdf) {
    content.push({ type: "document", source: { type: "base64", media_type: "application/pdf", data: base64Data } });
  } else {
    try { content.push({ type: "text", text: atob(base64Data).substring(0, 8000) }); } catch { content.push({ type: "text", text: "Could not decode file content" }); }
  }
  content.push({ type: "text", text: `Extract ALL information from this document/image about an educational activity, workshop, lecture, training, seminar, or plan. Return ONLY valid JSON, no markdown, no backticks, no explanation. Use these known fields when they match: title, type (one of: Lecture, Workshop, Training, Awareness Lecture, Seminar, Conference, Joint Program), date (YYYY-MM-DD — IMPORTANT: The academic year runs from September 2025 to August 2026. Activities in September, October, November, December must use year 2025. Activities in January through August must use year 2026. If only a month is mentioned without a year, apply this rule), presenter, delivery (In-Person/Online/Hybrid), target (Students/Faculty/Staff/External Audience/All), venue, duration, objectives, notes, department, source (College Plan/College Request/Department Initiative/University Initiative/Ministry Directive). For ANY other information found, add entries in a "customFields" array: [{"name":"...","value":"..."}]. If a field is not found leave it as "". Always include customFields. IMPORTANT: Preserve the ORIGINAL language of the text exactly as it appears in the document. If the text is Arabic, keep it Arabic. If English, keep it English. If mixed (e.g. Arabic body with English headers or names), keep both languages as they appear. Do NOT translate. For custom field names, use the same language as the document.` });
  try {
    const resp = await fetch("https://api.anthropic.com/v1/messages", {
      method: "POST",
      headers: makeApiHeaders(),
      body: JSON.stringify({
        model: "claude-sonnet-4-20250514",
        max_tokens: 4096,
        system: "You are a data extraction assistant. You MUST respond with ONLY a valid JSON object. No markdown, no backticks, no explanation text before or after. Just the raw JSON.",
        messages: [{ role: "user", content }]
      })
    });
    if (!resp.ok) { console.error("API error:", resp.status, await resp.text()); return null; }
    const data = await resp.json();
    if (data.error) { console.error("API returned error:", data.error); return null; }
    const text = (data.content || []).filter(b => b.type === "text").map(b => b.text).join("");
    if (!text) { console.error("Empty response from API"); return null; }
    let cleaned = text.replace(/```json\s*/gi, '').replace(/```\s*/g, '').trim();
    // Fix truncated response
    const ob = (cleaned.match(/\{/g)||[]).length, cb = (cleaned.match(/\}/g)||[]).length;
    if (ob > cb) { const lc = cleaned.lastIndexOf('}'); if(lc>0){ cleaned=cleaned.substring(0,lc+1); for(let i=0;i<ob-cb-1;i++)cleaned+='}'; }}
    try { return JSON.parse(cleaned); } catch(pe) {
      const m = cleaned.match(/\{[\s\S]*\}/);
      if (m) try { return JSON.parse(m[0]); } catch {}
      console.error("JSON parse failed:", pe, "Cleaned:", cleaned.substring(0, 300));
      return null;
    }
  } catch (e) { console.error("Extraction network error:", e); return null; }
}

async function extractBulkActivities(base64Data, mediaType, lang) {
  const isImage = mediaType.startsWith("image/");
  const isPdf = mediaType === "application/pdf";
  const content = [];
  if (isImage) {
    content.push({ type: "image", source: { type: "base64", media_type: mediaType, data: base64Data } });
  } else if (isPdf) {
    content.push({ type: "document", source: { type: "base64", media_type: "application/pdf", data: base64Data } });
  } else {
    try { content.push({ type: "text", text: atob(base64Data).substring(0, 12000) }); } catch { content.push({ type: "text", text: "Could not decode file content" }); }
  }
  content.push({ type: "text", text: `This is a yearly training/education plan document for academic year 2025-2026 (September 2025 to August 2026). Extract ALL activities listed. Return ONLY valid JSON, no markdown, no backticks. Format: {"activities": [{"title":"...","type":"...","date":"YYYY-MM-DD","presenter":"...","delivery":"...","target":"...","venue":"...","duration":"...","objectives":"...","notes":"...","department":"...","customFields":[]},...]}. CRITICAL DATE RULE: Activities in September, October, November, December must use year 2025. Activities in January through August must use year 2026. If only a month is mentioned without a year, apply this rule. Type must be one of: Lecture, Workshop, Training, Awareness Lecture, Seminar, Conference, Joint Program. Extract EVERY activity even if partial. Keep field values concise — short phrases, not full paragraphs. IMPORTANT: Preserve the ORIGINAL language of the text exactly as it appears in the document. If the text is Arabic, keep it Arabic. If English, keep it English. If mixed (e.g. Arabic body with English headers or names), keep both languages as they appear. Do NOT translate. For custom field names, use the same language as the document.` });
  try {
    const resp = await fetch("https://api.anthropic.com/v1/messages", {
      method: "POST",
      headers: makeApiHeaders(),
      body: JSON.stringify({
        model: "claude-sonnet-4-20250514",
        max_tokens: 8192,
        system: "You are a JSON-only data extractor. NEVER use markdown. NEVER use backticks. Output raw JSON and nothing else. Keep each activity description brief to fit within token limits.",
        messages: [{ role: "user", content }]
      })
    });
    if (!resp.ok) { console.error("Bulk API error:", resp.status, await resp.text()); return null; }
    const data = await resp.json();
    if (data.error) { console.error("Bulk API returned error:", data.error); return null; }
    const text = (data.content || []).filter(b => b.type === "text").map(b => b.text).join("");
    if (!text) { console.error("Empty bulk response"); return null; }
    // Clean markdown fences and extract JSON
    let cleaned = text.replace(/```json\s*/gi, '').replace(/```\s*/g, '').trim();
    // If response was truncated (no closing brackets), try to fix it
    const openBraces = (cleaned.match(/\{/g) || []).length;
    const closeBraces = (cleaned.match(/\}/g) || []).length;
    const openBrackets = (cleaned.match(/\[/g) || []).length;
    const closeBrackets = (cleaned.match(/\]/g) || []).length;
    // Close any unclosed strings, arrays, objects
    if (openBraces > closeBraces || openBrackets > closeBrackets) {
      // Find last complete activity object
      const lastComplete = cleaned.lastIndexOf('}');
      if (lastComplete > 0) {
        cleaned = cleaned.substring(0, lastComplete + 1);
        // Close remaining brackets
        const remaining = (cleaned.match(/\[/g) || []).length - (cleaned.match(/\]/g) || []).length;
        for (let i = 0; i < remaining; i++) cleaned += ']';
        const remainingBraces = (cleaned.match(/\{/g) || []).length - (cleaned.match(/\}/g) || []).length;
        for (let i = 0; i < remainingBraces; i++) cleaned += '}';
      }
    }
    let result;
    try { result = JSON.parse(cleaned); } catch(e1) {
      // Try to find JSON object in the text
      const m = cleaned.match(/\{[\s\S]*\}/);
      if (m) try { result = JSON.parse(m[0]); } catch(e2) {
        // Last resort: try to extract activities array manually
        const arrMatch = cleaned.match(/\[[\s\S]*\]/);
        if (arrMatch) try { result = { activities: JSON.parse(arrMatch[0]) }; } catch {}
      }
    }
    if (result && result.activities) return result.activities;
    if (Array.isArray(result)) return result;
    console.error("Bulk parse failed. Cleaned:", cleaned.substring(0, 500));
    return null;
  } catch (e) { console.error("Bulk extraction network error:", e); return null; }
}

async function extractCorrespondence(base64Data, mediaType, lang) {
  const isImage = mediaType.startsWith("image/");
  const isPdf = mediaType === "application/pdf";
  const content = [];
  if (isImage) {
    content.push({ type: "image", source: { type: "base64", media_type: mediaType, data: base64Data } });
  } else if (isPdf) {
    content.push({ type: "document", source: { type: "base64", media_type: "application/pdf", data: base64Data } });
  } else {
    try { content.push({ type: "text", text: atob(base64Data).substring(0, 8000) }); } catch { content.push({ type: "text", text: "Could not decode" }); }
  }
  content.push({ type: "text", text: `Extract information from this official correspondence/letter/memo/document. Return ONLY valid JSON, no markdown, no backticks. Fields: {"type":"Incoming or Outgoing","refNumber":"reference number if found","date":"YYYY-MM-DD","from":"sender name or organization","to":"recipient name or organization","subject":"subject or title of the document","notes":"any additional notes or summary of content","status":"Received or Sent","customFields":[{"name":"field name","value":"field value"}]}. Put any extra information found (like urgency, cc recipients, attachments mentioned, action required, deadline, department, signatures) into customFields. If a field is not found leave as "". IMPORTANT: Preserve the ORIGINAL language of the text exactly as it appears in the document. If the text is Arabic, keep it Arabic. If English, keep it English. If mixed (e.g. Arabic body with English headers or names), keep both languages as they appear. Do NOT translate. For custom field names, use the same language as the document.` });
  try {
    const resp = await fetch("https://api.anthropic.com/v1/messages", {
      method: "POST", headers: makeApiHeaders(),
      body: JSON.stringify({ model: "claude-sonnet-4-20250514", max_tokens: 2000, system: "You are a JSON-only data extractor. NEVER use markdown. NEVER use backticks. Output raw JSON only.", messages: [{ role: "user", content }] })
    });
    if (!resp.ok) { console.error("Corr API error:", resp.status); return null; }
    const data = await resp.json();
    if (data.error) { console.error("Corr API error:", data.error); return null; }
    const text = (data.content || []).filter(b => b.type === "text").map(b => b.text).join("");
    if (!text) return null;
    let cleaned = text.replace(/```json\s*/gi, '').replace(/```\s*/g, '').trim();
    const ob = (cleaned.match(/\{/g)||[]).length, cb = (cleaned.match(/\}/g)||[]).length;
    if (ob > cb) { const lc = cleaned.lastIndexOf('}'); if(lc>0){ cleaned=cleaned.substring(0,lc+1); for(let i=0;i<ob-cb-1;i++)cleaned+='}'; }}
    try { return JSON.parse(cleaned); } catch { const m = cleaned.match(/\{[\s\S]*\}/); if (m) try { return JSON.parse(m[0]); } catch {} return null; }
  } catch (e) { console.error("Corr extraction error:", e); return null; }
}

async function extractBulkCorrespondence(base64Data, mediaType, lang) {
  const content = [];
  if (mediaType.startsWith("image/")) content.push({ type: "image", source: { type: "base64", media_type: mediaType, data: base64Data } });
  else if (mediaType === "application/pdf") content.push({ type: "document", source: { type: "base64", media_type: "application/pdf", data: base64Data } });
  else { try { content.push({ type: "text", text: atob(base64Data).substring(0, 12000) }); } catch { content.push({ type: "text", text: "Could not decode" }); } }
  content.push({ type: "text", text: `This document contains one or more official letters, memos, or correspondence items. Extract ALL items. Return ONLY valid JSON, no markdown. Format: {"items":[{"type":"Incoming or Outgoing","refNumber":"","date":"YYYY-MM-DD","from":"","to":"","subject":"","notes":"","status":"Received","customFields":[{"name":"","value":""}]},...]}. Extract EVERY correspondence item. IMPORTANT: Preserve the ORIGINAL language of the text exactly as it appears in the document. If the text is Arabic, keep it Arabic. If English, keep it English. If mixed (e.g. Arabic body with English headers or names), keep both languages as they appear. Do NOT translate. For custom field names, use the same language as the document.` });
  try {
    const resp = await fetch("https://api.anthropic.com/v1/messages", { method: "POST", headers: makeApiHeaders(), body: JSON.stringify({ model: "claude-sonnet-4-20250514", max_tokens: 8192, system: "JSON-only extractor. No markdown. No backticks. Raw JSON only.", messages: [{ role: "user", content }] }) });
    if (!resp.ok) return null;
    const data = await resp.json(); if (data.error) return null;
    const text = (data.content || []).filter(b => b.type === "text").map(b => b.text).join("");
    let cleaned = text.replace(/```json\s*/gi, '').replace(/```\s*/g, '').trim();
    const ob=(cleaned.match(/\{/g)||[]).length,cb=(cleaned.match(/\}/g)||[]).length;
    if(ob>cb){const lc=cleaned.lastIndexOf('}');if(lc>0){cleaned=cleaned.substring(0,lc+1);for(let i=0;i<ob-cb-1;i++)cleaned+='}';const olb=(cleaned.match(/\[/g)||[]).length,clb=(cleaned.match(/\]/g)||[]).length;for(let i=0;i<olb-clb;i++)cleaned+=']';}}
    let result; try{result=JSON.parse(cleaned);}catch{const m=cleaned.match(/\{[\s\S]*\}/);if(m)try{result=JSON.parse(m[0]);}catch{}}
    if(result?.items)return result.items; if(Array.isArray(result))return result; return null;
  } catch { return null; }
}

async function extractTrainerFromFile(base64Data, mediaType, lang) {
  const content = [];
  if (mediaType.startsWith("image/")) content.push({ type: "image", source: { type: "base64", media_type: mediaType, data: base64Data } });
  else if (mediaType === "application/pdf") content.push({ type: "document", source: { type: "base64", media_type: "application/pdf", data: base64Data } });
  else { try { content.push({ type: "text", text: atob(base64Data).substring(0, 8000) }); } catch { content.push({ type: "text", text: "Could not decode" }); } }
  content.push({ type: "text", text: `Extract trainer/lecturer profile information from this document (CV, nomination form, credentials). Return ONLY valid JSON, no markdown. Fields: {"name":"full name","dob":"YYYY-MM-DD","gender":"Male or Female","phone":"","email":"","workplace":"","degree":"PhD/MSc/BSc/Board Certified","generalSpec":"general specialization","specificSpec":"specific specialization","internal":true or false,"notes":"","customFields":[{"name":"","value":""}]}. Put extra info (publications, experience years, certifications, languages, courses taught) into customFields. IMPORTANT: Preserve the ORIGINAL language of the text exactly as it appears in the document. If the text is Arabic, keep it Arabic. If English, keep it English. If mixed (e.g. Arabic body with English headers or names), keep both languages as they appear. Do NOT translate. For custom field names, use the same language as the document.` });
  try {
    const resp = await fetch("https://api.anthropic.com/v1/messages", { method: "POST", headers: makeApiHeaders(), body: JSON.stringify({ model: "claude-sonnet-4-20250514", max_tokens: 2000, system: "JSON-only extractor. No markdown. No backticks. Raw JSON only.", messages: [{ role: "user", content }] }) });
    if (!resp.ok) return null;
    const data = await resp.json(); if (data.error) return null;
    const text = (data.content || []).filter(b => b.type === "text").map(b => b.text).join("");
    let cleaned = text.replace(/```json\s*/gi, '').replace(/```\s*/g, '').trim();
    try { return JSON.parse(cleaned); } catch { const m = cleaned.match(/\{[\s\S]*\}/); if (m) try { return JSON.parse(m[0]); } catch {} return null; }
  } catch { return null; }
}

async function extractBulkTrainers(base64Data, mediaType, lang) {
  const content = [];
  if (mediaType.startsWith("image/")) content.push({ type: "image", source: { type: "base64", media_type: mediaType, data: base64Data } });
  else if (mediaType === "application/pdf") content.push({ type: "document", source: { type: "base64", media_type: "application/pdf", data: base64Data } });
  else { try { content.push({ type: "text", text: atob(base64Data).substring(0, 12000) }); } catch { content.push({ type: "text", text: "Could not decode" }); } }
  content.push({ type: "text", text: `This document contains a list of trainers, lecturers, or staff members. Extract ALL people. Return ONLY valid JSON. Format: {"trainers":[{"name":"","dob":"","gender":"Male or Female","phone":"","email":"","workplace":"","degree":"","generalSpec":"","specificSpec":"","internal":true,"notes":"","customFields":[]},...]}. Extract EVERY person even if partial. IMPORTANT: Preserve the ORIGINAL language of the text exactly as it appears in the document. If the text is Arabic, keep it Arabic. If English, keep it English. If mixed (e.g. Arabic body with English headers or names), keep both languages as they appear. Do NOT translate. For custom field names, use the same language as the document.` });
  try {
    const resp = await fetch("https://api.anthropic.com/v1/messages", { method: "POST", headers: makeApiHeaders(), body: JSON.stringify({ model: "claude-sonnet-4-20250514", max_tokens: 8192, system: "JSON-only extractor. No markdown. No backticks. Raw JSON only.", messages: [{ role: "user", content }] }) });
    if (!resp.ok) return null;
    const data = await resp.json(); if (data.error) return null;
    const text = (data.content || []).filter(b => b.type === "text").map(b => b.text).join("");
    let cleaned = text.replace(/```json\s*/gi, '').replace(/```\s*/g, '').trim();
    const ob=(cleaned.match(/\{/g)||[]).length,cb=(cleaned.match(/\}/g)||[]).length;
    if(ob>cb){const lc=cleaned.lastIndexOf('}');if(lc>0){cleaned=cleaned.substring(0,lc+1);for(let i=0;i<ob-cb-1;i++)cleaned+='}';const olb=(cleaned.match(/\[/g)||[]).length,clb=(cleaned.match(/\]/g)||[]).length;for(let i=0;i<olb-clb;i++)cleaned+=']';}}
    let result; try{result=JSON.parse(cleaned);}catch{const m=cleaned.match(/\{[\s\S]*\}/);if(m)try{result=JSON.parse(m[0]);}catch{}}
    if(result?.trainers)return result.trainers; if(Array.isArray(result))return result; return null;
  } catch { return null; }
}

async function extractEventReport(base64Data, mediaType, lang) {
  const content = [];
  if (mediaType.startsWith("image/")) content.push({ type: "image", source: { type: "base64", media_type: mediaType, data: base64Data } });
  else if (mediaType === "application/pdf") content.push({ type: "document", source: { type: "base64", media_type: "application/pdf", data: base64Data } });
  else { try { content.push({ type: "text", text: atob(base64Data).substring(0, 8000) }); } catch { content.push({ type: "text", text: "Could not decode" }); } }
  content.push({ type: "text", text: `Extract event report information from this document/image. This is a report about an educational activity, workshop, lecture, or training that has been conducted. Return ONLY valid JSON: {"activityTitle":"title of the activity","activityDate":"YYYY-MM-DD","activityType":"Lecture/Workshop/Training/Seminar/Conference","college":"college name","presenter":"presenter/lecturer name","totalParticipants":number,"participantNames":"comma-separated list of participant names if available","summary":"brief summary of the event","mediaNotified":true/false,"mediaLinks":"any media/publication links"}. IMPORTANT: Preserve the ORIGINAL language exactly as it appears. Do NOT translate.` });
  try {
    const resp = await fetch("https://api.anthropic.com/v1/messages", { method: "POST", headers: makeApiHeaders(), body: JSON.stringify({ model: "claude-sonnet-4-20250514", max_tokens: 4096, system: "JSON-only extractor. No markdown. No backticks. Raw JSON only.", messages: [{ role: "user", content }] }) });
    if (!resp.ok) return null;
    const data = await resp.json(); if (data.error) return null;
    const text = (data.content || []).filter(b => b.type === "text").map(b => b.text).join("");
    let cleaned = text.replace(/```json\s*/gi, '').replace(/```\s*/g, '').trim();
    const ob=(cleaned.match(/\{/g)||[]).length,cb=(cleaned.match(/\}/g)||[]).length;
    if(ob>cb){const lc=cleaned.lastIndexOf('}');if(lc>0){cleaned=cleaned.substring(0,lc+1);for(let i=0;i<ob-cb-1;i++)cleaned+='}';}}
    try { return JSON.parse(cleaned); } catch { const m = cleaned.match(/\{[\s\S]*\}/); if (m) try { return JSON.parse(m[0]); } catch {} return null; }
  } catch { return null; }
}

function gid(){return "id_"+Date.now()+"_"+Math.random().toString(36).substr(2,5);}
const sC=s=>({"Planned":"#3498db","Approved":"#1abc9c","In Progress":"#f39c12","Completed":"#27ae60","Postponed":"#95a5a6","Cancelled":"#e74c3c","Pending Review":"#9b59b6"}[s]||"#888");

// ─── STYLES ───────────────────────────────────────────────────────────
const lbl={fontSize:12,fontWeight:500,color:"var(--color-text-secondary)",display:"block",marginBottom:4};
const inp={width:"100%",padding:"8px 12px",border:"1px solid var(--color-border-tertiary)",borderRadius:8,fontSize:13,boxSizing:"border-box",background:"#ffffff",color:"var(--color-text-primary)"};
const btnP={padding:"8px 18px",background:"#1a4f72",color:"#fff",border:"none",borderRadius:8,fontSize:13,fontWeight:600,cursor:"pointer"};
const btnC={padding:"8px 18px",background:"transparent",border:"1px solid var(--color-border-secondary)",borderRadius:8,fontSize:13,cursor:"pointer",color:"var(--color-text-secondary)"};
const btnI={padding:6,background:"none",border:"1px solid var(--color-border-tertiary)",borderRadius:6,cursor:"pointer",display:"flex",alignItems:"center",justifyContent:"center"};
const crd={background:"#ffffff",borderRadius:12,padding:"18px 22px",border:"0.5px solid var(--color-border-tertiary)"};

const Ic=({n,s=18,c="currentColor"})=>{const d={dashboard:<><rect x="3" y="3" width="7" height="7" rx="1"/><rect x="14" y="3" width="7" height="7" rx="1"/><rect x="3" y="14" width="7" height="7" rx="1"/><rect x="14" y="14" width="7" height="7" rx="1"/></>,plan:<><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><path d="M14 2v6h6"/><line x1="16" y1="13" x2="8" y2="13"/><line x1="16" y1="17" x2="8" y2="17"/></>,activity:<><circle cx="12" cy="12" r="10"/><polyline points="12 6 12 12 16 14"/></>,star:<><polygon points="12 2 15.09 8.26 22 9.27 17 14.14 18.18 21.02 12 17.77 5.82 21.02 7 14.14 2 9.27 8.91 8.26 12 2"/></>,users:<><path d="M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2"/><circle cx="9" cy="7" r="4"/><path d="M23 21v-2a4 4 0 0 0-3-3.87"/></>,logout:<><path d="M9 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h4"/><polyline points="16 17 21 12 16 7"/><line x1="21" y1="12" x2="9" y2="12"/></>,plus:<><line x1="12" y1="5" x2="12" y2="19"/><line x1="5" y1="12" x2="19" y2="12"/></>,edit:<><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></>,trash:<><polyline points="3 6 5 6 21 6"/><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/></>,check:<><polyline points="20 6 9 17 4 12"/></>,search:<><circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/></>,mail:<><path d="M4 4h16c1.1 0 2 .9 2 2v12c0 1.1-.9 2-2 2H4c-1.1 0-2-.9-2-2V6c0-1.1.9-2 2-2z"/><polyline points="22,6 12,13 2,6"/></>,report:<><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><path d="M14 2v6h6"/></>,download:<><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></>,backup:<><polyline points="1 4 1 10 7 10"/><path d="M3.51 15a9 9 0 1 0 2.13-9.36L1 10"/></>,lock:<><rect x="3" y="11" width="18" height="11" rx="2"/><path d="M7 11V7a5 5 0 0 1 10 0v4"/></>,user:<><path d="M20 21v-2a4 4 0 0 0-4-4H8a4 4 0 0 0-4 4v2"/><circle cx="12" cy="7" r="4"/></>,settings:<><circle cx="12" cy="12" r="3"/><path d="M19.4 15a1.65 1.65 0 0 0 .33 1.82l.06.06a2 2 0 0 1 0 2.83 2 2 0 0 1-2.83 0l-.06-.06a1.65 1.65 0 0 0-1.82-.33 1.65 1.65 0 0 0-1 1.51V21a2 2 0 0 1-2 2 2 2 0 0 1-2-2v-.09A1.65 1.65 0 0 0 9 19.4a1.65 1.65 0 0 0-1.82.33l-.06.06a2 2 0 0 1-2.83 0 2 2 0 0 1 0-2.83l.06-.06A1.65 1.65 0 0 0 4.68 15a1.65 1.65 0 0 0-1.51-1H3a2 2 0 0 1-2-2 2 2 0 0 1 2-2h.09A1.65 1.65 0 0 0 4.6 9a1.65 1.65 0 0 0-.33-1.82l-.06-.06a2 2 0 0 1 0-2.83 2 2 0 0 1 2.83 0l.06.06A1.65 1.65 0 0 0 9 4.68a1.65 1.65 0 0 0 1-1.51V3a2 2 0 0 1 2-2 2 2 0 0 1 2 2v.09a1.65 1.65 0 0 0 1 1.51 1.65 1.65 0 0 0 1.82-.33l.06-.06a2 2 0 0 1 2.83 0 2 2 0 0 1 0 2.83l-.06.06a1.65 1.65 0 0 0-.33 1.82V9a1.65 1.65 0 0 0 1.51 1H21a2 2 0 0 1 2 2 2 2 0 0 1-2 2h-.09a1.65 1.65 0 0 0-1.51 1z"/></>,bell:<><path d="M18 8A6 6 0 0 0 6 8c0 7-3 9-3 9h18s-3-2-3-9"/><path d="M13.73 21a2 2 0 0 1-3.46 0"/></>,bar:<><line x1="18" y1="20" x2="18" y2="10"/><line x1="12" y1="20" x2="12" y2="4"/><line x1="6" y1="20" x2="6" y2="14"/></>,printer:<><polyline points="6 9 6 2 18 2 18 9"/><path d="M6 18H4a2 2 0 0 1-2-2v-5a2 2 0 0 1 2-2h16a2 2 0 0 1 2 2v5a2 2 0 0 1-2 2h-2"/><rect x="6" y="14" width="12" height="8"/></>};return <svg width={s} height={s} viewBox="0 0 24 24" fill="none" stroke={c} strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">{d[n]}</svg>;};

// ─── APP ──────────────────────────────────────────────────────────────
export default function App(){
  const [user,setUser]=useState(null);
  const [page,setPage]=useState("dashboard");
  const [plans,setPlans]=useState(SAMPLE_PLANS);
  const [inits,setInits]=useState(SAMPLE_INIT);
  const [trainers,setTrainers]=useState(SAMPLE_TRAINERS);
  const [corr,setCorr]=useState(SAMPLE_CORR);
  const [reports,setReports]=useState([]);
  const [staffRoster,setStaffRoster]=useState([]);
  const [staffMerges,setStaffMerges]=useState([]);
  const [directives,setDirectives]=useState([]);
  const [users,setUsers]=useState(DEFAULT_USERS);
  const [config,setConfig]=useState({lang:"en",platformName:"",platformSub:"",customLogo:""});
  const [selPlan,setSelPlan]=useState(null);
  const [sidebar,setSidebar]=useState(true);
  const [isMobile,setIsMobile]=useState(false);
  useEffect(()=>{
    const check=()=>{const m=window.innerWidth<768;setIsMobile(m);if(m)setSidebar(false);};
    check();window.addEventListener('resize',check);return()=>window.removeEventListener('resize',check);
  },[]);
  const [modal,setModal]=useState(null);
  const [loading,setLoading]=useState(true);
  const [showNotif,setShowNotif]=useState(false);
  const [stickyNotes,setStickyNotes]=useState([]);
  const [auditLog,setAuditLog]=useState([]);
  const [securityDismissed,setSecurityDismissed]=useState(false);
  const [lastActivity,setLastActivity]=useState(Date.now());

  const t=T[config.lang]||T.en;
  const isRtl=t.dir==="rtl";

  // Session timeout - check every 60s
  useEffect(()=>{
    if(!user)return;
    const interval=setInterval(()=>{
      if(Date.now()-lastActivity>SESSION_TIMEOUT_MS){
        setUser(null);setPage("dashboard");setSelPlan(null);
        pS(K.session,null).catch(()=>{});
        addAuditEntry(t.sessionTimeout||"Session expired",user,"Auto-logout after 30 min inactivity");
      }
    },60000);
    return ()=>clearInterval(interval);
  },[user,lastActivity]);

  // Reset activity timer on any interaction
  useEffect(()=>{
    const reset=()=>setLastActivity(Date.now());
    window.addEventListener("click",reset);
    window.addEventListener("keydown",reset);
    return ()=>{window.removeEventListener("click",reset);window.removeEventListener("keydown",reset);};
  },[]);

  useEffect(()=>{(async()=>{
    try{localStorage.removeItem("ced-session");}catch{}
    try{localStorage.removeItem("ced2-session");}catch{}
    try{localStorage.removeItem("ced3-session");}catch{}
    const s=await pG(K.session,null);
    const storedUsers=await sG(K.users,DEFAULT_USERS);
    setUsers(storedUsers);
    // Validate session and check timeout
    if(s?.user?.username&&s?.loginTime){
      const elapsed=Date.now()-s.loginTime;
      if(elapsed<SESSION_TIMEOUT_MS){
        const v=storedUsers.find(u=>u.username===s.user.username);
        if(v)setUser(s.user);
      }
    }
    setPlans(await sG(K.plans,SAMPLE_PLANS));
    setInits(await sG(K.init,SAMPLE_INIT));
    setTrainers(await sG(K.trainers,SAMPLE_TRAINERS));
    setCorr(await sG(K.corr,SAMPLE_CORR));
    setReports(await sG(K.reports,[]));
    setStaffRoster(await sG(K.roster,[]));
    setStaffMerges(await sG(K.staffMerges,[]));
    setDirectives(await sG(K.directives,[]));
    setConfig(await sG(K.config,{lang:"en",platformName:"",platformSub:"",customLogo:""}));
    setStickyNotes(await sG(K.notes,[]));
    setAuditLog(await sG(K.audit,[]));
    const sd=await pG("cedp-security-dismissed",false);
    setSecurityDismissed(sd);
    setLoading(false);
  })();},[]);

  const sv=(k,setter)=>async v=>{setter(v);await sS(k,v);};
  const savePlans=sv(K.plans,setPlans),saveInits=sv(K.init,setInits),saveTrainers=sv(K.trainers,setTrainers),saveCorr=sv(K.corr,setCorr),saveReports=sv(K.reports,setReports),saveRoster=sv(K.roster,setStaffRoster),saveStaffMerges=sv(K.staffMerges,setStaffMerges),saveDirectives=sv(K.directives,setDirectives),saveUsers=sv(K.users,setUsers),saveConfig=sv(K.config,setConfig),saveNotes=sv(K.notes,setStickyNotes);

  const audit=async(action,details)=>{const updated=await addAuditEntry(action,user,details);setAuditLog(updated);};

  const login=async(u,p)=>{
    const uLower=u.trim().toLowerCase();
    const found=users.find(x=>x.username===uLower);
    if(!found){await addAuditEntry("Login failed",{name:u,role:""},`Unknown username: ${u}`);return false;}
    // Support both hashed and plain text passwords (migration)
    let valid=false;
    if(found.passwordHash){valid=await verifyPassword(p,found.passwordHash);}
    else if(found.password){
      valid=(found.password===p);
      // Migrate to hashed password on successful plain text login
      if(valid){const hash=await hashPassword(p);const updated=users.map(x=>x.id===found.id?{...x,passwordHash:hash,password:undefined}:x);await saveUsers(updated);}
    }
    if(valid){
      setUser(found);setLastActivity(Date.now());
      await pS(K.session,{user:found,loginTime:Date.now()});
      await addAuditEntry("Login successful",found,"");
      return true;
    }
    await addAuditEntry("Login failed",{name:u,role:""},`Wrong password for: ${u}`);
    return false;
  };

  const logout=async()=>{
    await audit("Logout","");
    setUser(null);setPage("dashboard");setSelPlan(null);
    try{localStorage.removeItem(K.session);}catch{}
  };

  const [syncing,setSyncing]=useState(false);
  const [lastSync,setLastSync]=useState(Date.now());
  const syncData=async()=>{
    setSyncing(true);
    try{
      setPlans(await sG(K.plans,SAMPLE_PLANS));
      setInits(await sG(K.init,SAMPLE_INIT));
      setTrainers(await sG(K.trainers,SAMPLE_TRAINERS));
      setCorr(await sG(K.corr,SAMPLE_CORR));
      setReports(await sG(K.reports,[]));
      setStaffRoster(await sG(K.roster,[]));
    setStaffMerges(await sG(K.staffMerges,[]));
      setDirectives(await sG(K.directives,[]));
      setConfig(await sG(K.config,{lang:"en",platformName:"",platformSub:"",customLogo:""}));
      setStickyNotes(await sG(K.notes,[]));
      setAuditLog(await sG(K.audit,[]));
      setLastSync(Date.now());
    }catch(e){console.error("Sync error:",e);}
    setSyncing(false);
  };

  // Auto-sync every 2 minutes to keep data fresh across tabs/devices
  useEffect(()=>{
    if(!user)return;
    const interval=setInterval(()=>syncData(),120000);
    return ()=>clearInterval(interval);
  },[user]);

  const isAdmin=user?.role==="admin";
  const canEdit=user&&(user.role==="admin"||user.role==="college");
  const pName=config.platformName||t.appDefaultName;
  const pSub=config.platformSub||t.appDefaultSub;
  const logoSrc=config.customLogo||LOGO;
  const COLLEGES=config.colleges&&config.colleges.length>0?config.colleges:DEFAULT_COLLEGES;

  const allActs=useMemo(()=>[
    ...plans.flatMap(p=>p.activities.map(a=>({...a,college:p.college,planId:p.id}))),
    ...inits.map(i=>({...i,planId:null})),
  ],[plans,inits]);

  const upcoming7=useMemo(()=>{
    const now=new Date();const week=new Date(now.getTime()+7*86400000);
    return allActs.filter(a=>a.date&&new Date(a.date)>=now&&new Date(a.date)<=week&&a.status!=="Completed"&&a.status!=="Cancelled"&&a.status!=="Postponed").sort((a,b)=>new Date(a.date)-new Date(b.date));
  },[allActs]);

  if(loading)return <div style={{height:"100vh",display:"flex",alignItems:"center",justifyContent:"center",fontFamily:t.font}}>Loading...</div>;
  if(users.length===0)return <SetupScreen t={t} saveUsers={saveUsers} config={config} saveConfig={saveConfig}/>;
  if(!user)return <Login onLogin={login} t={t} pName={pName} pSub={pSub} logoSrc={logoSrc} users={users} saveUsers={saveUsers}/>;

  const navItems=[
    {id:"dashboard",icon:"dashboard",label:t.dashboard},
    {id:"plans",icon:"plan",label:t.yearlyPlans},
    {id:"activities",icon:"activity",label:t.allActivities},
    {id:"initiatives",icon:"star",label:t.initiatives},
    {id:"directives",icon:"mail",label:t.directivesPage||"Directives"},
    {id:"trainers",icon:"user",label:t.trainers},
    {id:"staff",icon:"users",label:t.staff||"College Staff"},
    {id:"reports",icon:"report",label:t.reports},
    {id:"correspondence",icon:"mail",label:t.correspondence},
    {id:"analytics",icon:"bar",label:t.analytics},
    {id:"backup",icon:"backup",label:t.backup},
  ];
  if(isAdmin)navItems.push({id:"audit",icon:"lock",label:t.auditLog},{id:"settings",icon:"settings",label:t.settings},{id:"access",icon:"lock",label:t.access});

  return(
    <div dir={t.dir} style={{display:"flex",height:"100vh",fontFamily:t.font,fontSize:14,color:"var(--color-text-primary)",background:"var(--color-background-tertiary)",overflow:"hidden"}}>
      <style>{`.sidebar-scroll::-webkit-scrollbar{width:5px}.sidebar-scroll::-webkit-scrollbar-track{background:transparent}.sidebar-scroll::-webkit-scrollbar-thumb{background:rgba(255,255,255,0.15);border-radius:4px}.sidebar-scroll::-webkit-scrollbar-thumb:hover{background:rgba(255,255,255,0.25)}`}</style>
      {/* MOBILE BACKDROP */}
      {isMobile&&sidebar&&<div onClick={()=>setSidebar(false)} style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.5)",zIndex:998}}/>}
      {/* SIDEBAR */}
      <div style={{width:sidebar?230:isMobile?0:56,background:"#0d3249",display:"flex",flexDirection:"column",flexShrink:0,transition:"width 0.2s",overflow:"hidden",order:isRtl?1:0,...(isMobile?{position:"fixed",top:0,[isRtl?"right":"left"]:0,height:"100vh",zIndex:999,width:sidebar?260:0}:{})}}>
        <div style={{padding:sidebar?"12px 12px 8px":"8px 8px",display:"flex",alignItems:"center",gap:8,borderBottom:"1px solid rgba(255,255,255,0.07)",cursor:"pointer",justifyContent:sidebar?"flex-start":"center"}} onClick={()=>setSidebar(!sidebar)}>
          <img src={logoSrc} alt="" style={{width:38,height:38,borderRadius:"50%",objectFit:"contain",background:"#0d3249",flexShrink:0}}/>
          {sidebar&&<div style={{minWidth:0,flex:1}}><div style={{fontSize:11,fontWeight:700,color:"#c9a84c",lineHeight:1.2,whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis"}}>{pName}</div><div style={{fontSize:9,color:"#8fa7bf",marginTop:1,whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis"}}>{pSub}</div></div>}
          {sidebar&&isMobile&&<span style={{color:"#8fa7bf",fontSize:18,cursor:"pointer"}}>✕</span>}
        </div>
        <div className="sidebar-scroll" style={{padding:"6px 4px",flex:1,overflowY:"auto"}}>
          {navItems.map(it=>(
            <button key={it.id} onClick={()=>{setPage(it.id);setSelPlan(null);if(isMobile)setSidebar(false);}} title={it.label} style={{width:"100%",display:"flex",alignItems:"center",gap:8,padding:sidebar?"8px 10px":"8px 0",borderRadius:7,border:"none",cursor:"pointer",marginBottom:1,background:page===it.id?"rgba(201,168,76,0.15)":"transparent",justifyContent:sidebar?"flex-start":"center"}}>
              <Ic n={it.icon} s={17} c={page===it.id?"#c9a84c":"#8fa7bf"}/>
              {sidebar&&<span style={{fontSize:12.5,fontWeight:page===it.id?600:400,color:page===it.id?"#c9a84c":"#b0c4d8",whiteSpace:"nowrap"}}>{it.label}</span>}
            </button>
          ))}
        </div>
        <div style={{padding:sidebar?"10px 12px":"10px 4px",borderTop:"1px solid rgba(255,255,255,0.07)",display:"flex",alignItems:"center",gap:6,justifyContent:sidebar?"flex-start":"center"}}>
          <div style={{width:28,height:28,borderRadius:"50%",background:"rgba(201,168,76,0.2)",display:"flex",alignItems:"center",justifyContent:"center",fontSize:11,fontWeight:600,color:"#c9a84c",flexShrink:0}}>{user.name[0]}</div>
          {sidebar&&<div style={{flex:1,minWidth:0}}><div style={{fontSize:11,fontWeight:600,color:"#d4e6f1",overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{user.name}</div><div style={{fontSize:9,color:"#6c8ca6",textTransform:"capitalize"}}>{user.role==="college"?t.collegeRep:user.role==="admin"?t.admin:t.viewer}</div></div>}
          <button onClick={logout} style={{background:"none",border:"none",cursor:"pointer",padding:2}} title={t.logout}><Ic n="logout" s={15} c="#8fa7bf"/></button>
        </div>
      </div>

      {/* MAIN */}
      <div style={{flex:1,display:"flex",flexDirection:"column",overflow:"hidden",order:isRtl?0:1}}>
        <div style={{height:48,padding:"0 14px",display:"flex",alignItems:"center",gap:10,background:"#ffffff",borderBottom:"0.5px solid var(--color-border-tertiary)",flexShrink:0}}>
          {/* Hamburger for mobile */}
          {isMobile&&<button onClick={()=>setSidebar(true)} style={{background:"none",border:"none",cursor:"pointer",padding:4,fontSize:20,lineHeight:1}}>☰</button>}
          {isMobile&&<span style={{fontSize:12,fontWeight:600,color:"#1a5276",flex:1,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{navItems.find(n=>n.id===page)?.label||""}</span>}
          {!isMobile&&<div style={{flex:1}}/>}
          {/* Sync button */}
          <button onClick={syncData} disabled={syncing} title={syncing?"Syncing...":"Sync data from storage"} style={{background:"none",border:"none",cursor:"pointer",padding:4,opacity:syncing?0.3:0.7,transition:"opacity 0.2s"}}>
            <svg width={16} height={16} viewBox="0 0 24 24" fill="none" stroke="#8fa7bf" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round"><polyline points="23 4 23 10 17 10"/><polyline points="1 20 1 14 7 14"/><path d="M3.51 9a9 9 0 0 1 14.85-3.36L23 10M1 14l4.64 4.36A9 9 0 0 0 20.49 15"/></svg>
          </button>
          {/* Notification bell */}
          <div style={{position:"relative"}}>
            <button onClick={()=>setShowNotif(!showNotif)} style={{background:"none",border:"none",cursor:"pointer",padding:4,position:"relative"}}>
              <svg width={18} height={18} viewBox="0 0 24 24" fill="none" stroke="#8fa7bf" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round"><path d="M18 8A6 6 0 0 0 6 8c0 7-3 9-3 9h18s-3-2-3-9"/><path d="M13.73 21a2 2 0 0 1-3.46 0"/></svg>
              {upcoming7.length>0&&<span style={{position:"absolute",top:0,right:0,width:16,height:16,borderRadius:"50%",background:"#e74c3c",color:"#fff",fontSize:9,display:"flex",alignItems:"center",justifyContent:"center",fontWeight:700}}>{upcoming7.length}</span>}
            </button>
            {showNotif&&<div style={{position:"absolute",[isRtl?"left":"right"]:0,top:"100%",marginTop:4,width:300,background:"#ffffff",borderRadius:10,border:"0.5px solid var(--color-border-tertiary)",boxShadow:"0 4px 20px rgba(0,0,0,0.1)",zIndex:100,padding:12,maxHeight:300,overflowY:"auto"}}>
              <div style={{fontSize:12,fontWeight:600,marginBottom:8}}>{t.notifications}</div>
              {upcoming7.length===0&&<div style={{fontSize:11,color:"var(--color-text-tertiary)"}}>{t.noNotif}</div>}
              {upcoming7.map(a=><div key={a.id} style={{padding:"6px 0",borderBottom:"0.5px solid var(--color-border-tertiary)",fontSize:11}}>
                <div style={{fontWeight:600}}>{a.title}</div>
                <div style={{color:"var(--color-text-secondary)"}}>📅 {a.date} · {a.type}</div>
              </div>)}
            </div>}
          </div>
          {/* Language toggle */}
          <button onClick={()=>{const nl=config.lang==="en"?"ar":"en";saveConfig({...config,lang:nl});}} style={{padding:"4px 10px",borderRadius:16,border:"1px solid var(--color-border-tertiary)",background:"transparent",fontSize:11,cursor:"pointer",color:"var(--color-text-secondary)",fontWeight:600}}>{config.lang==="en"?"عربي":"EN"}</button>
          <span style={{fontSize:10,padding:"3px 10px",borderRadius:20,background:isAdmin?"#e8d5a3":user.role==="viewer"?"var(--color-background-info)":"var(--color-background-success)",color:isAdmin?"#7d6119":user.role==="viewer"?"var(--color-text-info)":"var(--color-text-success)",fontWeight:600,textTransform:"capitalize"}}>{user.role==="college"?t.collegeRep:user.role==="admin"?t.admin:t.viewer}</span>
        </div>
        <div style={{flex:1,padding:isMobile?"10px 10px":"14px 18px",overflow:"auto"}} onClick={()=>showNotif&&setShowNotif(false)}>
          {page==="dashboard"&&<Dashboard t={t} plans={plans} inits={inits} allActs={allActs} reports={reports} isRtl={isRtl} COLLEGES={COLLEGES} isMobile={isMobile}/>}
          {page==="plans"&&!selPlan&&<PlansPage t={t} plans={plans} savePlans={savePlans} user={user} canEdit={canEdit} isAdmin={isAdmin} onSelect={setSelPlan} COLLEGES={COLLEGES}/>}
          {page==="plans"&&selPlan&&<PlanDetail t={t} plan={selPlan} plans={plans} savePlans={savePlans} canEdit={canEdit} isAdmin={isAdmin} user={user} onBack={()=>setSelPlan(null)} trainers={trainers} logoSrc={logoSrc} pName={pName} pSub={pSub}/>}
          {page==="activities"&&<ActivitiesPage t={t} allActs={allActs} logoSrc={logoSrc} pName={pName} pSub={pSub} COLLEGES={COLLEGES}/>}
          {page==="initiatives"&&<InitiativesPage t={t} inits={inits} saveInits={saveInits} isAdmin={isAdmin} trainers={trainers} logoSrc={logoSrc} pName={pName} pSub={pSub} COLLEGES={COLLEGES}/>}
          {page==="directives"&&<DirectivesPage t={t} directives={directives} saveDirectives={saveDirectives} allActs={allActs} inits={inits} saveInits={saveInits} plans={plans} savePlans={savePlans} COLLEGES={COLLEGES} isAdmin={isAdmin} logoSrc={logoSrc} pName={pName}/>}
          {page==="trainers"&&<TrainersPage t={t} trainers={trainers} saveTrainers={saveTrainers} isAdmin={isAdmin} canEdit={canEdit} COLLEGES={COLLEGES} allActs={allActs} logoSrc={logoSrc} pName={pName}/>}
          {page==="staff"&&<StaffPage t={t} allActs={allActs} corr={corr} trainers={trainers} saveTrainers={saveTrainers} COLLEGES={COLLEGES} isAdmin={isAdmin} staffRoster={staffRoster} saveRoster={saveRoster} staffMerges={staffMerges} saveStaffMerges={saveStaffMerges} logoSrc={logoSrc} pName={pName} reports={reports}/>}
          {page==="reports"&&<ReportsPage t={t} reports={reports} saveReports={saveReports} allActs={allActs} isAdmin={isAdmin} canEdit={canEdit} logoSrc={logoSrc} pName={pName} COLLEGES={COLLEGES} plans={plans} savePlans={savePlans} inits={inits} saveInits={saveInits}/>}
          {page==="correspondence"&&<CorrPage t={t} corr={corr} saveCorr={saveCorr} isAdmin={isAdmin} canEdit={canEdit} COLLEGES={COLLEGES}/>}
          {page==="analytics"&&<AnalyticsPage t={t} allActs={allActs} plans={plans} isRtl={isRtl} COLLEGES={COLLEGES} reports={reports} logoSrc={logoSrc} pName={pName}/>}
          {page==="backup"&&<BackupPage t={t} plans={plans} inits={inits} trainers={trainers} corr={corr} reports={reports} users={users} config={config} notes={stickyNotes} savePlans={savePlans} saveInits={saveInits} saveTrainers={saveTrainers} saveCorr={saveCorr} saveReports={saveReports} saveUsers={saveUsers} saveConfig={saveConfig} saveNotes={saveNotes}/>}
          {page==="settings"&&isAdmin&&<SettingsPage t={t} config={config} saveConfig={saveConfig} users={users} saveUsers={saveUsers} currentUser={user} audit={audit} COLLEGES={COLLEGES} plans={plans} savePlans={savePlans} inits={inits} saveInits={saveInits} trainers={trainers} saveTrainers={saveTrainers} corr={corr} saveCorr={saveCorr} reports={reports} saveReports={saveReports}/>}
          {page==="audit"&&isAdmin&&<AuditPage t={t} auditLog={auditLog}/>}
          {page==="access"&&isAdmin&&<AccessPage t={t}/>}
          {/* Security notice banner */}
          {user&&!securityDismissed&&<div style={{position:"fixed",bottom:70,[isRtl?"left":"right"]:80,width:340,background:"#fef9e7",border:"1px solid #f0c674",borderRadius:12,padding:16,boxShadow:"0 4px 20px rgba(0,0,0,0.1)",zIndex:80}}>
            <div style={{display:"flex",gap:8,marginBottom:8}}><span style={{fontSize:18}}>🔒</span><span style={{fontSize:13,fontWeight:700,color:"#7d6608"}}>{t.securityNotice}</span></div>
            <p style={{fontSize:11,color:"#5d4e0e",lineHeight:1.7,margin:"0 0 10px"}}>{t.securityMsg}</p>
            <button onClick={()=>{setSecurityDismissed(true);pS("cedp-security-dismissed",true);}} style={{...btnP,padding:"6px 16px",fontSize:11,background:"#7d6608"}}>{t.dismissNotice}</button>
          </div>}
        </div>
      </div>
      {modal&&<div onClick={()=>setModal(null)} style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.45)",display:"flex",alignItems:"center",justifyContent:"center",zIndex:999}}><div onClick={e=>e.stopPropagation()} style={{background:"#ffffff",borderRadius:14,padding:24,maxWidth:600,width:"92%",maxHeight:"80vh",overflow:"auto"}}>{modal}<button onClick={()=>setModal(null)} style={{...btnC,marginTop:14}}>{t.cancel}</button></div></div>}
      <StickyNotesPanel t={t} notes={stickyNotes} saveNotes={saveNotes} isRtl={isRtl}/>
      <HelpPanel page={page} isRtl={isRtl} lang={config.lang}/>
    </div>
  );
}

// ─── FIRST-TIME SETUP ─────────────────────────────────────────────────
function SetupScreen({t,saveUsers,config,saveConfig}){
  const [name,setName]=useState("");
  const [username,setUsername]=useState("");
  const [pw,setPw]=useState("");
  const [pw2,setPw2]=useState("");
  const [err,setErr]=useState("");
  const [busy,setBusy]=useState(false);
  const [lang,setLang]=useState(config.lang||"en");
  const [recoveryKey,setRecoveryKey]=useState("");
  const [step,setStep]=useState(1); // 1=form, 2=show recovery key

  const curT=T[lang]||T.en;

  const doSetup=async()=>{
    setErr("");
    if(!name.trim()||!username.trim()||!pw.trim()){setErr(curT.setupError);return;}
    if(pw.length<6){setErr(curT.setupPasswordShort);return;}
    if(pw!==pw2){setErr(curT.setupPasswordMatch);return;}
    setBusy(true);
    try{
      const hash=await hashPassword(pw);
      const admin={id:gid(),username:username.trim().toLowerCase(),passwordHash:hash,role:"admin",name:name.trim(),college:null};
      await saveUsers([admin]);
      // Generate recovery key automatically
      const key=generateRecoveryKey();
      await saveRecoveryKeyHash(key);
      setRecoveryKey(key);
      // Save language preference
      await saveConfig({...config,lang});
      // Log setup
      await addAuditEntry("System setup complete",admin,"Admin account created");
      setStep(2);
    }catch(e){setErr("Error: "+e.message);}
    setBusy(false);
  };

  if(step===2){
    return(
      <div dir={curT.dir} style={{minHeight:"100vh",display:"flex",alignItems:"center",justifyContent:"center",background:"#f5f5f5",fontFamily:"system-ui,sans-serif",padding:20}}>
        <div style={{width:"100%",maxWidth:440,background:"#fff",borderRadius:8,padding:"32px 28px",border:"1px solid #ddd"}}>
          <div style={{marginBottom:20}}>
            <h2 style={{fontSize:16,fontWeight:600,color:"#333",margin:"0 0 6px"}}>Recovery key</h2>
            <p style={{fontSize:12,color:"#888",margin:0,lineHeight:1.6}}>Save this key. You will need it to reset your password.</p>
          </div>
          <div style={{background:"#f9f9f9",border:"1px solid #e0e0e0",borderRadius:6,padding:20,textAlign:"center",marginBottom:16}}>
            <div style={{fontFamily:"monospace",fontSize:22,fontWeight:700,color:"#333",letterSpacing:2,userSelect:"all"}}>{recoveryKey}</div>
          </div>
          <div style={{display:"flex",gap:8,justifyContent:"center",marginBottom:16}}>
            <button onClick={()=>{navigator.clipboard?.writeText(recoveryKey);}} style={{padding:"6px 16px",border:"1px solid #ccc",borderRadius:4,background:"#fff",cursor:"pointer",fontSize:12}}>Copy</button>
          </div>
          <button onClick={()=>setTimeout(()=>window.location.reload(),500)} style={{width:"100%",padding:12,background:"#333",color:"#fff",border:"none",borderRadius:4,fontSize:14,cursor:"pointer"}}>Continue</button>
        </div>
      </div>
    );
  }

  return(
    <div dir={curT.dir} style={{minHeight:"100vh",display:"flex",alignItems:"center",justifyContent:"center",background:"#f5f5f5",fontFamily:"system-ui,sans-serif",padding:20}}>
      <div style={{width:"100%",maxWidth:400,background:"#fff",borderRadius:8,padding:"32px 28px",border:"1px solid #ddd"}}>
        <div style={{marginBottom:24}}>
          <h1 style={{fontSize:16,fontWeight:600,color:"#333",margin:"0 0 4px"}}>Account setup</h1>
          <p style={{fontSize:12,color:"#999",margin:0}}>Create your account to continue.</p>
        </div>

        {/* Language selector - subtle */}
        <div style={{display:"flex",gap:6,marginBottom:18}}>
          <button onClick={()=>setLang("en")} style={{padding:"4px 12px",borderRadius:4,border:lang==="en"?"1px solid #333":"1px solid #ddd",background:lang==="en"?"#f0f0f0":"transparent",cursor:"pointer",fontSize:11,color:"#555"}}>EN</button>
          <button onClick={()=>setLang("ar")} style={{padding:"4px 12px",borderRadius:4,border:lang==="ar"?"1px solid #333":"1px solid #ddd",background:lang==="ar"?"#f0f0f0":"transparent",cursor:"pointer",fontSize:11,color:"#555"}}>AR</button>
        </div>

        {err&&<div style={{background:"#fee",color:"#c00",padding:"8px 10px",borderRadius:4,fontSize:12,marginBottom:12}}>{err}</div>}

        <div style={{marginBottom:12}}>
          <label style={{fontSize:11,color:"#777",display:"block",marginBottom:3}}>Name</label>
          <input value={name} onChange={e=>setName(e.target.value)} style={{width:"100%",padding:"8px 10px",border:"1px solid #ddd",borderRadius:4,fontSize:13,boxSizing:"border-box"}}/>
        </div>
        <div style={{marginBottom:12}}>
          <label style={{fontSize:11,color:"#777",display:"block",marginBottom:3}}>Username</label>
          <input value={username} onChange={e=>setUsername(e.target.value)} style={{width:"100%",padding:"8px 10px",border:"1px solid #ddd",borderRadius:4,fontSize:13,boxSizing:"border-box"}}/>
        </div>
        <div style={{marginBottom:12}}>
          <label style={{fontSize:11,color:"#777",display:"block",marginBottom:3}}>Password</label>
          <input type="password" value={pw} onChange={e=>setPw(e.target.value)} style={{width:"100%",padding:"8px 10px",border:"1px solid #ddd",borderRadius:4,fontSize:13,boxSizing:"border-box"}}/>
        </div>
        <div style={{marginBottom:20}}>
          <label style={{fontSize:11,color:"#777",display:"block",marginBottom:3}}>Confirm password</label>
          <input type="password" value={pw2} onChange={e=>setPw2(e.target.value)} onKeyDown={e=>e.key==="Enter"&&doSetup()} style={{width:"100%",padding:"8px 10px",border:"1px solid #ddd",borderRadius:4,fontSize:13,boxSizing:"border-box"}}/>
        </div>

        <button onClick={doSetup} disabled={busy} style={{width:"100%",padding:10,background:"#333",color:"#fff",border:"none",borderRadius:4,fontSize:13,cursor:"pointer",opacity:busy?0.6:1}}>{busy?"...":"Create account"}</button>
      </div>
    </div>
  );
}

// ─── LOGIN ────────────────────────────────────────────────────────────
function Login({onLogin,t,pName,pSub,logoSrc,users,saveUsers}){
  const [u,sU]=useState("");const [p,sP]=useState("");const [e,sE]=useState("");const [busy,sB]=useState(false);
  const [showRecovery,setShowRecovery]=useState(false);
  const [recUser,setRecUser]=useState("");
  const [recKey,setRecKey]=useState("");
  const [recNewPw,setRecNewPw]=useState("");
  const [recMsg,setRecMsg]=useState("");
  const [recBusy,setRecBusy]=useState(false);

  const go=async()=>{if(busy||!u.trim()||!p.trim())return;sE("");sB(true);try{if(!await onLogin(u.trim(),p.trim()))sE(t.invalidCred);}catch{sE("Error");}sB(false);};

  const doRecover=async()=>{
    if(!recUser||!recKey.trim()||!recNewPw.trim()){setRecMsg("Fill all fields");return;}
    if(recNewPw.trim().length<4){setRecMsg("Password too short (min 4)");return;}
    setRecBusy(true);setRecMsg("");
    const valid=await verifyRecoveryKey(recKey.trim());
    if(!valid){setRecMsg(t.recoveryFailed);setRecBusy(false);await addAuditEntry("Recovery failed",{name:recUser,role:""},"Invalid recovery key");return;}
    const hash=await hashPassword(recNewPw.trim());
    const updated=users.map(x=>x.username===recUser?{...x,passwordHash:hash,password:undefined}:x);
    await saveUsers(updated);
    await addAuditEntry("Password reset via recovery",{name:recUser,role:""},"Recovery key used");
    setRecMsg(t.recoverySuccess);
    setRecBusy(false);
    setTimeout(()=>{setShowRecovery(false);sU(recUser);},2000);
  };

  if(showRecovery){
    return(
      <div dir={t.dir} style={{minHeight:"100vh",display:"flex",alignItems:"center",justifyContent:"center",background:"#f5f5f5",fontFamily:"system-ui,sans-serif",padding:20}}>
        <div style={{width:"100%",maxWidth:400,background:"#fff",borderRadius:8,padding:"32px 28px",border:"1px solid #ddd"}}>
          <div style={{marginBottom:20}}>
            <h2 style={{fontSize:16,fontWeight:600,color:"#333",margin:"0 0 4px"}}>Reset password</h2>
            <p style={{fontSize:12,color:"#999",margin:0}}>Enter your recovery key to reset.</p>
          </div>
          {recMsg&&<div style={{background:recMsg===t.recoverySuccess?"#d5f5e3":"#fee",color:recMsg===t.recoverySuccess?"#196f3d":"#c00",padding:"8px 10px",borderRadius:4,fontSize:12,marginBottom:12}}>{recMsg}</div>}
          <div style={{marginBottom:12}}>
            <label style={{fontSize:11,color:"#777",display:"block",marginBottom:3}}>Username</label>
            <select value={recUser} onChange={v=>setRecUser(v.target.value)} style={{width:"100%",padding:"8px 10px",border:"1px solid #ddd",borderRadius:4,fontSize:13,boxSizing:"border-box"}}>
              <option value="">--</option>
              {users.map(x=><option key={x.id} value={x.username}>{x.username}</option>)}
            </select>
          </div>
          <div style={{marginBottom:12}}>
            <label style={{fontSize:11,color:"#777",display:"block",marginBottom:3}}>Recovery key</label>
            <input value={recKey} onChange={v=>setRecKey(v.target.value)} placeholder="XXXXX-XXXXX-XXXXX-XXXXX" style={{width:"100%",padding:"8px 10px",border:"1px solid #ddd",borderRadius:4,fontSize:13,boxSizing:"border-box",fontFamily:"monospace",letterSpacing:1}}/>
          </div>
          <div style={{marginBottom:20}}>
            <label style={{fontSize:11,color:"#777",display:"block",marginBottom:3}}>New password</label>
            <input type="password" value={recNewPw} onChange={v=>setRecNewPw(v.target.value)} style={{width:"100%",padding:"8px 10px",border:"1px solid #ddd",borderRadius:4,fontSize:13,boxSizing:"border-box"}}/>
          </div>
          <button onClick={doRecover} disabled={recBusy} style={{width:"100%",padding:10,background:"#333",color:"#fff",border:"none",borderRadius:4,fontSize:13,cursor:"pointer",opacity:recBusy?0.6:1}}>{recBusy?"...":"Reset password"}</button>
          <button onClick={()=>setShowRecovery(false)} style={{width:"100%",marginTop:8,padding:8,background:"transparent",border:"1px solid #ddd",borderRadius:4,fontSize:12,cursor:"pointer",color:"#666"}}>Back</button>
        </div>
      </div>
    );
  }

  const isBranded = !!(pName && pName !== "Continual Education Department" && pName !== "شعبة التعليم المستمر") || !!(logoSrc && logoSrc !== LOGO);

  return(
    <div dir={t.dir} style={{minHeight:"100vh",display:"flex",alignItems:"center",justifyContent:"center",background:isBranded?"linear-gradient(135deg,#0b2d42 0%,#1a5276 50%,#1e6f91 100%)":"#f5f5f5",fontFamily:isBranded?t.font:"system-ui,sans-serif",padding:20}}>
      <div style={{width:"100%",maxWidth:isBranded?400:380,background:isBranded?"rgba(255,255,255,0.97)":"#fff",borderRadius:isBranded?16:8,padding:isBranded?"36px 32px":"32px 28px",boxShadow:isBranded?"0 20px 60px rgba(0,0,0,0.3)":"none",border:isBranded?"none":"1px solid #ddd"}}>
        {isBranded ? (
          <div style={{textAlign:"center",marginBottom:28}}>
            <img src={logoSrc} alt="" style={{width:90,height:90,borderRadius:"50%",marginBottom:12,border:"3px solid #c9a84c",objectFit:"contain",background:"#0d3249"}}/>
            <h1 style={{fontSize:18,fontWeight:700,color:"#0e3554",margin:0}}>{pName}</h1>
            <p style={{fontSize:12,color:"#6c7a89",margin:"4px 0 0"}}>{pSub}</p>
          </div>
        ) : (
          <div style={{marginBottom:24}}>
            <h1 style={{fontSize:16,fontWeight:600,color:"#333",margin:"0 0 4px"}}>Sign in</h1>
            <p style={{fontSize:12,color:"#999",margin:0}}>Enter your credentials to continue.</p>
          </div>
        )}
        {e&&<div style={{background:isBranded?"#fdecea":"#fee",color:isBranded?"#c0392b":"#c00",padding:"8px 10px",borderRadius:isBranded?8:4,fontSize:12,marginBottom:12}}>{e}</div>}
        <div style={{marginBottom:12}}>
          <label style={{fontSize:11,color:isBranded?"#5d6d7e":"#777",display:"block",marginBottom:3}}>{isBranded?t.username:"Username"}</label>
          <input value={u} onChange={v=>sU(v.target.value)} onKeyDown={v=>v.key==="Enter"&&go()} style={{width:"100%",padding:"8px 10px",border:`1px solid ${isBranded?"#d5dbdb":"#ddd"}`,borderRadius:isBranded?8:4,fontSize:13,boxSizing:"border-box"}}/>
        </div>
        <div style={{marginBottom:20}}>
          <label style={{fontSize:11,color:isBranded?"#5d6d7e":"#777",display:"block",marginBottom:3}}>{isBranded?t.password:"Password"}</label>
          <input type="password" value={p} onChange={v=>sP(v.target.value)} onKeyDown={v=>v.key==="Enter"&&go()} style={{width:"100%",padding:"8px 10px",border:`1px solid ${isBranded?"#d5dbdb":"#ddd"}`,borderRadius:isBranded?8:4,fontSize:13,boxSizing:"border-box"}}/>
        </div>
        <button onClick={go} disabled={busy} style={{width:"100%",padding:isBranded?12:10,background:isBranded?"#1a4f72":"#333",color:"#fff",border:"none",borderRadius:isBranded?8:4,fontSize:isBranded?15:13,fontWeight:isBranded?600:400,cursor:"pointer",opacity:busy?0.6:1}}>{busy?"...":(isBranded?t.login:"Sign in")}</button>
        <div style={{textAlign:"center",marginTop:12}}>
          <button onClick={()=>setShowRecovery(true)} style={{background:"none",border:"none",color:isBranded?"#2980b9":"#888",cursor:"pointer",fontSize:11,textDecoration:isBranded?"underline":"none"}}>{isBranded?t.forgotPassword:"Forgot password?"}</button>
        </div>
      </div>
    </div>
  );
}

// ─── STAT CARD ────────────────────────────────────────────────────────
function SC({label,value,sub,color,onClick,active}){return <div onClick={onClick} style={{...crd,cursor:onClick?"pointer":"default",border:active?"2px solid "+(color||"#1a5276"):"1px solid var(--color-border-tertiary)",transition:"all 0.15s"}}><div style={{fontSize:11,color:"var(--color-text-secondary)",marginBottom:4,fontWeight:500}}>{label}</div><div style={{fontSize:26,fontWeight:700,color:color||"var(--color-text-primary)",lineHeight:1}}>{value}</div>{sub&&<div style={{fontSize:10,color:"var(--color-text-tertiary)",marginTop:5}}>{sub}</div>}</div>;}

// ─── DASHBOARD ────────────────────────────────────────────────────────
function Dashboard({t,plans,inits,allActs,reports,isRtl,COLLEGES,isMobile}){
  const [fCollege,setFCollege]=useState("All");
  const [drillDown,setDrillDown]=useState(null); // null, "total", "completed", "overdue", "attendees", "initiatives", "reports"
  const acts=fCollege==="All"?allActs:allActs.filter(a=>a.college===fCollege);
  const fPlans=fCollege==="All"?plans:plans.filter(p=>p.college===fCollege);
  const now=new Date();
  const comp=acts.filter(a=>a.status==="Completed");
  const overdue=acts.filter(a=>a.date&&new Date(a.date)<now&&a.status!=="Completed"&&a.status!=="Cancelled"&&a.status!=="Postponed");
  const totalAtt=acts.reduce((s,a)=>s+(a.attendees||0),0);
  const upcoming=acts.filter(a=>a.date&&new Date(a.date)>=now&&a.status!=="Completed"&&a.status!=="Cancelled"&&a.status!=="Postponed").sort((a,b)=>new Date(a.date)-new Date(b.date)).slice(0,6);
  const bySource={};acts.forEach(a=>{bySource[a.source]=(bySource[a.source]||0)+1;});
  const planned=acts.filter(a=>a.status==="Planned");
  const approved=acts.filter(a=>a.status==="Approved");
  const cancelled=acts.filter(a=>a.status==="Cancelled");
  const postponed=acts.filter(a=>a.status==="Postponed");
  const inProgress=acts.filter(a=>a.status==="In Progress");
  // Media stats
  const rptMap={};(reports||[]).forEach(r=>{rptMap[r.activityTitle]=r;});
  const mediaPublished=acts.filter(a=>{const rpt=rptMap[a.title]||{};return a.mediaStatus==="Published"||(rpt.mediaNotified&&rpt.mediaLinks);}).length;
  const mediaContacted=acts.filter(a=>{const rpt=rptMap[a.title]||{};const s=a.mediaStatus||"Not contacted";return !((s==="Published")||(rpt.mediaNotified&&rpt.mediaLinks))&&(s==="Contacted"||rpt.mediaNotified);}).length;
  const MN=isRtl?MONTHS_AR:MONTHS_EN;
  const byMonth=Array(12).fill(0);acts.forEach(a=>{if(a.date){const m=new Date(a.date).getMonth();if(!isNaN(m))byMonth[MONTH_MAP[m]]++;}});
  const monthData=byMonth.map((v,i)=>({name:MN[i],value:v}));
  const pieData=Object.entries(bySource).map(([k,v])=>({name:k,value:v}));

  const toggle=k=>setDrillDown(drillDown===k?null:k);

  // Drill-down data
  const drillData={
    total:acts,
    completed:comp,
    overdue:overdue,
    attendees:acts.filter(a=>(a.attendees||0)>0).sort((a,b)=>(b.attendees||0)-(a.attendees||0)),
    initiatives:fCollege==="All"?inits:inits.filter(i=>i.college===fCollege),
    reports:fCollege==="All"?reports:reports.filter(r=>r.college===fCollege),
    planned,approved,cancelled,postponed,inProgress,
  };

  const renderDrillList=(items,type)=>{
    if(!items||items.length===0)return <p style={{fontSize:12,color:"var(--color-text-tertiary)",padding:10}}>No items</p>;
    return <div style={{maxHeight:250,overflowY:"auto"}}>
      {items.slice(0,50).map((a,i)=>(
        <div key={a.id||i} style={{display:"flex",alignItems:"center",gap:8,padding:"6px 8px",borderBottom:"0.5px solid var(--color-border-tertiary)",fontSize:11}}>
          <div style={{width:5,height:5,borderRadius:"50%",background:sC(a.status),flexShrink:0}}/>
          <div style={{flex:1,minWidth:0}}>
            <div style={{fontWeight:500,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{a.title||a.activityTitle||a.subject||"-"}</div>
            <div style={{fontSize:10,color:"var(--color-text-tertiary)"}}>
              {a.date||a.activityDate||""}{a.college?` · ${a.college.replace("College of ","").replace("Department of ","")}`:""}{a.type||a.activityType?` · ${a.type||a.activityType}`:""}{type==="attendees"?` · 👥 ${a.attendees||0}`:""}{type==="overdue"?` · ⚠️ ${Math.ceil((now-new Date(a.date))/86400000)} days`:""}{a.presenter?` · ${a.presenter}`:""}
            </div>
          </div>
          <span style={{fontSize:9,padding:"2px 6px",borderRadius:8,background:`${sC(a.status)}18`,color:sC(a.status),fontWeight:600,flexShrink:0}}>{a.status||""}</span>
        </div>
      ))}
      {items.length>50&&<div style={{fontSize:10,color:"var(--color-text-tertiary)",padding:8}}>...and {items.length-50} more</div>}
    </div>;
  };

  return(
    <div>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:16}}>
        <div><h2 style={{fontSize:18,fontWeight:700,margin:"0 0 3px"}}>{t.dashboard}</h2><p style={{fontSize:12,color:"var(--color-text-secondary)",margin:0}}>{t.academicYear} · <span style={{fontSize:10,color:"var(--color-text-tertiary)"}}>{t.clickToView}</span></p></div>
        <select value={fCollege} onChange={e=>{setFCollege(e.target.value);setDrillDown(null);}} style={{...inp,width:"auto",minWidth:200,padding:"6px 10px",fontSize:12}}>
          <option value="All">🏛️ {t.all} — {allActs.length} {t.acts}</option>
          {(COLLEGES||[]).map(c=><option key={c} value={c}>{c.replace("College of ","").replace("Department of ","")} — {allActs.filter(a=>a.college===c).length} {t.acts}</option>)}
        </select>
      </div>
      <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(130px,1fr))",gap:10,marginBottom:4}}>
        <SC label={t.totalActivities} value={acts.length} sub={`${fPlans.length} ${t.plansSubmitted}`} onClick={()=>toggle("total")} active={drillDown==="total"}/>
        <SC label={t.completed} value={comp.length} color="#27ae60" sub={`${acts.length?Math.round(comp.length/acts.length*100):0}% ${t.completionRate}`} onClick={()=>toggle("completed")} active={drillDown==="completed"}/>
        <SC label={"⚠️ "+t.overdue} value={overdue.length} color="#e74c3c" sub={t.overdueDesc} onClick={()=>toggle("overdue")} active={drillDown==="overdue"}/>
        <SC label={t.totalAttendees} value={totalAtt.toLocaleString()} color="#8e44ad" sub={t.attendanceSource} onClick={()=>toggle("attendees")} active={drillDown==="attendees"}/>
        <SC label={t.initiatives} value={(fCollege==="All"?inits:inits.filter(i=>i.college===fCollege)).length} color="#e67e22" onClick={()=>toggle("initiatives")} active={drillDown==="initiatives"}/>
        <SC label={t.reports} value={(fCollege==="All"?reports:reports.filter(r=>r.college===fCollege)).length} color="#2980b9" onClick={()=>toggle("reports")} active={drillDown==="reports"}/>
      </div>
      {/* ROW 2: Status breakdown + Media */}
      <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(110px,1fr))",gap:8,marginBottom:4,marginTop:8}}>
        <SC label="📝 Planned" value={planned.length} color="#3498db" onClick={()=>toggle("planned")} active={drillDown==="planned"}/>
        <SC label="✅ Approved" value={approved.length} color="#1abc9c" onClick={()=>toggle("approved")} active={drillDown==="approved"}/>
        <SC label="🔄 In Progress" value={inProgress.length} color="#f39c12" onClick={()=>toggle("inProgress")} active={drillDown==="inProgress"}/>
        <SC label="⏸️ Postponed" value={postponed.length} color="#95a5a6" onClick={()=>toggle("postponed")} active={drillDown==="postponed"}/>
        <SC label="🚫 Cancelled" value={cancelled.length} color="#7f8c8d" onClick={()=>toggle("cancelled")} active={drillDown==="cancelled"}/>
        <SC label="📰 Published" value={mediaPublished} color="#196f3d" sub={`${acts.length?Math.round(mediaPublished/acts.length*100):0}% coverage`}/>
        <SC label="📞 Media contacted" value={mediaContacted} color="#2471a3"/>
      </div>

      {/* DRILL-DOWN PANEL */}
      {drillDown&&<div style={{...crd,marginBottom:16,border:`2px solid ${drillDown==="overdue"?"#e74c3c":drillDown==="completed"?"#27ae60":"#1a5276"}`}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:8}}>
          <h3 style={{fontSize:13,fontWeight:600,margin:0}}>
            {drillDown==="total"&&`📊 ${t.totalActivities}`}
            {drillDown==="completed"&&`✅ ${t.completed}`}
            {drillDown==="overdue"&&`⚠️ ${t.overdue}`}
            {drillDown==="attendees"&&`👥 ${t.totalAttendees}`}
            {drillDown==="initiatives"&&`⭐ ${t.initiatives}`}
            {drillDown==="reports"&&`📋 ${t.reports}`}
            {drillDown==="planned"&&`📝 Planned`}
            {drillDown==="approved"&&`✅ Approved`}
            {drillDown==="cancelled"&&`🚫 Cancelled`}
            {drillDown==="postponed"&&`⏸️ Postponed`}
            {drillDown==="inProgress"&&`🔄 In Progress`}
            {" "}({(drillData[drillDown]||[]).length})
          </h3>
          <button onClick={()=>setDrillDown(null)} style={{...btnC,padding:"3px 10px",fontSize:10}}>{t.hideDetails} ✕</button>
        </div>
        {renderDrillList(drillData[drillDown],drillDown)}
      </div>}

      <div style={{display:"grid",gridTemplateColumns:isMobile?"1fr":"1fr 1fr",gap:14,marginBottom:20}}>
        <div style={crd}>
          <h3 style={{fontSize:13,fontWeight:600,margin:"0 0 10px"}}>{t.monthlyTrend}{fCollege!=="All"&&<span style={{fontWeight:400,fontSize:11,color:"var(--color-text-tertiary)"}}> — {fCollege.replace("College of ","").replace("Department of ","")}</span>}</h3>
          <div style={{height:180}}><ResponsiveContainer width="100%" height="100%"><BarChart data={monthData}><XAxis dataKey="name" tick={{fontSize:10}}/><YAxis allowDecimals={false} tick={{fontSize:10}}/><Tooltip/><Bar dataKey="value" fill="#1a5276" radius={[3,3,0,0]}/></BarChart></ResponsiveContainer></div>
        </div>
        <div style={crd}>
          <h3 style={{fontSize:13,fontWeight:600,margin:"0 0 10px"}}>{t.sourceDistribution}</h3>
          <div style={{height:180}}><ResponsiveContainer width="100%" height="100%"><PieChart><Pie data={pieData} cx="50%" cy="50%" outerRadius={65} dataKey="value" label={({name,percent})=>`${name.split(" ")[0]} ${Math.round(percent*100)}%`} labelLine={false} style={{fontSize:9}}>{pieData.map((e,i)=><Cell key={i} fill={PIE_COLORS[i%PIE_COLORS.length]}/>)}</Pie><Tooltip/></PieChart></ResponsiveContainer></div>
        </div>
      </div>
      {/* Media coverage overview */}
      {(()=>{
        const mPie=[
          {name:"Published",value:mediaPublished},
          {name:"Contacted",value:mediaContacted},
          {name:"Not contacted",value:Math.max(0,acts.length-mediaPublished-mediaContacted)},
        ].filter(d=>d.value>0);
        const MC=["#27ae60","#3498db","#95a5a6"];
        return mediaPublished+mediaContacted>0?<div style={{...crd,marginBottom:14}}>
          <h3 style={{fontSize:13,fontWeight:600,margin:"0 0 10px"}}>📰 Media Coverage</h3>
          <div style={{display:"grid",gridTemplateColumns:isMobile?"1fr":"1fr 1fr",gap:14,alignItems:"center"}}>
            <div style={{height:150}}><ResponsiveContainer width="100%" height="100%"><PieChart><Pie data={mPie} cx="50%" cy="50%" outerRadius={55} dataKey="value" label={({name,percent})=>`${name} ${Math.round(percent*100)}%`} labelLine={false} style={{fontSize:9}}>{mPie.map((e,i)=><Cell key={i} fill={MC[i%MC.length]}/>)}</Pie><Tooltip/></PieChart></ResponsiveContainer></div>
            <div style={{display:"flex",flexDirection:"column",gap:6}}>
              <div style={{fontSize:12}}>📰 <b style={{color:"#27ae60"}}>{mediaPublished}</b> Published ({acts.length?Math.round(mediaPublished/acts.length*100):0}% coverage)</div>
              <div style={{fontSize:12}}>📞 <b style={{color:"#3498db"}}>{mediaContacted}</b> Contacted</div>
              <div style={{fontSize:12}}>⬜ <b style={{color:"#95a5a6"}}>{Math.max(0,acts.length-mediaPublished-mediaContacted)}</b> Not covered</div>
            </div>
          </div>
        </div>:null;
      })()}
      <div style={crd}>
        <h3 style={{fontSize:13,fontWeight:600,margin:"0 0 10px"}}>{t.upcomingActivities}</h3>
        {upcoming.length===0&&<p style={{fontSize:12,color:"var(--color-text-tertiary)"}}>{t.noUpcoming}</p>}
        {upcoming.map(a=><div key={a.id} style={{display:"flex",alignItems:"center",gap:8,padding:"6px 0",borderBottom:"0.5px solid var(--color-border-tertiary)"}}>
          <div style={{width:5,height:5,borderRadius:"50%",background:sC(a.status),flexShrink:0}}/>
          <div style={{flex:1,minWidth:0}}><div style={{fontSize:12,fontWeight:500,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{a.title}</div><div style={{fontSize:10,color:"var(--color-text-tertiary)"}}>{a.date} · {a.type}{a.college&&` · ${a.college.replace("College of ","").replace("Department of ","")}`}</div></div>
          <span style={{fontSize:9,padding:"2px 7px",borderRadius:10,background:`${sC(a.status)}18`,color:sC(a.status),fontWeight:600}}>{a.status}</span>
        </div>)}
      </div>
    </div>
  );
}

// ─── ANALYTICS PAGE ───────────────────────────────────────────────────
function AnalyticsPage({t,allActs,plans,isRtl,COLLEGES,reports,logoSrc,pName}){
  const [fCollege,setFCollege]=useState("All");
  const [showExportPanel,setShowExportPanel]=useState(false);
  const [expSec,setExpSec]=useState({summary:true,monthly:true,byType:true,byStatus:true,collegeTable:true,completed:true,overdue:true,noReport:true,activityList:true,media:true});
  const togSec=k=>setExpSec({...expSec,[k]:!expSec[k]});
  const acts=fCollege==="All"?allActs:allActs.filter(a=>a.college===fCollege);
  const now=new Date();
  const MN=isRtl?MONTHS_AR:MONTHS_EN;
  const byMonth=Array(12).fill(0);const compByMonth=Array(12).fill(0);
  acts.forEach(a=>{if(a.date){const m=new Date(a.date).getMonth();if(!isNaN(m)){byMonth[MONTH_MAP[m]]++;if(a.status==="Completed")compByMonth[MONTH_MAP[m]]++;}}});
  const lineData=MN.map((n,i)=>({name:n,total:byMonth[i],completed:compByMonth[i]}));
  const byType={};acts.forEach(a=>{byType[a.type]=(byType[a.type]||0)+1;});
  const typeData=Object.entries(byType).map(([k,v])=>({name:k,value:v}));
  const byStatus={};acts.forEach(a=>{byStatus[a.status]=(byStatus[a.status]||0)+1;});
  const statusData=Object.entries(byStatus).map(([k,v])=>({name:k,value:v}));

  // College comparison data (includes BOTH plan activities AND initiatives)
  const reportTitles=new Set((reports||[]).map(r=>r.activityTitle));
  const rptMapA={};(reports||[]).forEach(r=>{rptMapA[r.activityTitle]=r;});
  const collegeStats=(COLLEGES||[]).map(c=>{
    const cActs=allActs.filter(a=>a.college===c);
    const fromPlans=cActs.filter(a=>a.planId);
    const fromInits=cActs.filter(a=>!a.planId);
    const completed=cActs.filter(a=>a.status==="Completed");
    const planned=cActs.filter(a=>a.status==="Planned").length;
    const approved=cActs.filter(a=>a.status==="Approved").length;
    const inProgress=cActs.filter(a=>a.status==="In Progress").length;
    const cancelled=cActs.filter(a=>a.status==="Cancelled").length;
    const postponed=cActs.filter(a=>a.status==="Postponed").length;
    const overdue=cActs.filter(a=>a.date&&new Date(a.date)<now&&a.status!=="Completed"&&a.status!=="Cancelled"&&a.status!=="Postponed");
    const noReport=completed.filter(a=>!reportTitles.has(a.title)&&!(a.participants&&a.participants.trim())&&!(a.attendees>0));
    const totalAtt=cActs.reduce((s,a)=>s+(a.attendees||0),0);
    const mediaPub=cActs.filter(a=>{const rpt=rptMapA[a.title]||{};return a.mediaStatus==="Published"||(rpt.mediaNotified&&rpt.mediaLinks);}).length;
    const short=c.replace("College of ","").replace("Department of ","");
    return{name:short,full:c,total:cActs.length,plans:fromPlans.length,inits:fromInits.length,completed:completed.length,planned,approved,inProgress,cancelled,postponed,overdue:overdue.length,noReport:noReport.length,rate:cActs.length?Math.round(completed.length/cActs.length*100):0,attendees:totalAtt,mediaPub};
  }).filter(c=>c.total>0).sort((a,b)=>b.total-a.total);

  const totalColData=collegeStats.map(c=>({name:c.name,total:c.total,plans:c.plans,inits:c.inits}));
  const compColData=collegeStats.map(c=>({name:c.name,completed:c.completed,overdue:c.overdue,noReport:c.noReport,cancelled:c.cancelled,postponed:c.postponed}));
  const rateColData=collegeStats.map(c=>({name:c.name,rate:c.rate}));
  const attColData=collegeStats.filter(c=>c.attendees>0).map(c=>({name:c.name,attendees:c.attendees}));

  // Export functions
  const exportAnalyticsExcel=()=>{
    const wb=XLSX.utils.book_new();
    if(expSec.collegeTable&&collegeStats.length>0){const csData=collegeStats.map(c=>({"College":c.full,"Total activities":c.total,"From plans":c.plans,"From initiatives":c.inits,"Planned":c.planned,"Approved":c.approved,"Completed":c.completed,"Completion %":c.rate+"%","Overdue":c.overdue,"Cancelled":c.cancelled,"Postponed":c.postponed,"No report filed":c.noReport,"Attendees":c.attendees,"Media published":c.mediaPub}));XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(csData),"College Summary");}
    if(expSec.monthly){const mData=lineData.map(m=>({"Month":m.name,"Total":m.total,"Completed":m.completed}));XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(mData),"Monthly Trend");}
    if(expSec.byType){const tData=typeData.map(x=>({"Type":x.name,"Count":x.value}));XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(tData),"By Type");}
    if(expSec.byStatus){const sData=statusData.map(x=>({"Status":x.name,"Count":x.value}));XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(sData),"By Status");}
    if(expSec.activityList){const aData=acts.map(a=>{const row={"Title":a.title,"Type":a.type,"Date":a.date,"Status":a.status,"Source":a.source,"College":a.college||"","Presenter":a.presenter||"","Delivery":a.delivery||"","Venue":a.venue||"","Attendees":a.attendees||0,"Objectives":a.objectives||"","Notes":a.notes||""};if(a.participants)row["Participants"]=a.participants;if(a.customFields)a.customFields.filter(f=>f.name).forEach(f=>{row[f.name]=f.value;});return row;});XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(aData),"Activities Detail");}
    if(expSec.completed){const ca=acts.filter(a=>a.status==="Completed");if(ca.length>0){const rm={};(reports||[]).forEach(r=>{rm[r.activityTitle]=r;});const cd=ca.map(a=>{const rpt=rm[a.title]||{};return{"Title":a.title,"Type":a.type,"Date":a.date,"College":a.college||"","Source":a.source||"","Presenter":a.presenter||rpt.presenter||"","Attendees":a.attendees||rpt.totalParticipants||0,"Participant names":a.participants||rpt.participantNames||"","Summary":rpt.summary||"","Media notified":rpt.mediaNotified?"✅":"❌","Media status":a.mediaStatus||rpt.mediaStatus||"","Media channel":a.mediaChannel||rpt.mediaChannel||"","Report filed":rpt.createdDate||"⚠️ No report"};});XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(cd),"Completed Events");}}
    if(expSec.overdue){const oa=acts.filter(a=>a.date&&new Date(a.date)<now&&a.status!=="Completed"&&a.status!=="Cancelled"&&a.status!=="Postponed");if(oa.length>0){const od=oa.map(a=>({"Title":a.title,"Date":a.date,"Days overdue":Math.ceil((now-new Date(a.date))/86400000),"Status":a.status,"College":a.college||"","Source":a.source||""}));XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(od),"Overdue");}}
    if(expSec.noReport){const nr=acts.filter(a=>a.status==="Completed"&&!reportTitles.has(a.title)&&!(a.participants&&a.participants.trim())&&!(a.attendees>0));if(nr.length>0){const nd=nr.map(a=>({"Title":a.title,"Date":a.date,"College":a.college||"","Source":a.source||""}));XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(nd),"No Report Filed");}}
    if(expSec.media){const rptM={};(reports||[]).forEach(r=>{rptM[r.activityTitle]=r;});const md=acts.filter(a=>{const rpt=rptM[a.title]||{};return a.mediaStatus==="Published"||a.mediaStatus==="Contacted"||rpt.mediaNotified||rpt.mediaLinks;}).map(a=>{const rpt=rptM[a.title]||{};return{"Title":a.title,"Date":a.date,"College":a.college||"","Status":a.mediaStatus||"","Channel":a.mediaChannel||"","Contacted date":a.mediaContactedDate||"","Published date":a.mediaPublishedDate||"","Media notified (report)":rpt.mediaNotified?"✅":"","Links":rpt.mediaLinks||"","Notes":a.mediaCoverageNotes||""};});if(md.length>0)XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(md),"Media Coverage");}
    if(wb.SheetNames.length===0)XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet([{"Note":"No sections selected"}]),"Info");
    const label=fCollege!=="All"?fCollege.replace(/College of |Department of /g,""):"All_Colleges";
    XLSX.writeFile(wb,`Analytics_${label}_${new Date().toISOString().split('T')[0]}.xlsx`);
  };

  const printAnalytics=()=>{
    const dir=isRtl?"rtl":"ltr";const al=isRtl?"right":"left";
    const label=fCollege!=="All"?fCollege:"All Colleges";
    const overdueActs=acts.filter(a=>a.date&&new Date(a.date)<now&&a.status!=="Completed"&&a.status!=="Cancelled"&&a.status!=="Postponed");
    const noRpt=acts.filter(a=>a.status==="Completed"&&!reportTitles.has(a.title)&&!(a.participants&&a.participants.trim())&&!(a.attendees>0));
    let html=`<html dir="${dir}"><head><meta charset="utf-8"><style>
      @import url('https://fonts.googleapis.com/css2?family=Noto+Sans+Arabic:wght@400;700&display=swap');
      body{font-family:'Noto Sans Arabic','Noto Sans',sans-serif;padding:30px;font-size:12px;color:#333;}
      h1{color:#1a5276;font-size:20px;margin:0 0 4px;} h2{color:#1a5276;font-size:15px;margin:20px 0 8px;border-bottom:2px solid #1a5276;padding-bottom:4px;}
      table{width:100%;border-collapse:collapse;margin:8px 0 16px;} th,td{border:1px solid #ddd;padding:5px 8px;text-align:${al};font-size:11px;} th{background:#f4f6f7;font-weight:700;}
      .badge{padding:2px 8px;border-radius:8px;font-size:10px;font-weight:600;} .g{background:#d5f5e3;color:#196f3d;} .y{background:#fef5e7;color:#b7950b;} .r{background:#fdedec;color:#c0392b;}
      .stat{display:inline-block;padding:8px 16px;margin:0 8px 8px 0;border-radius:8px;background:#f4f6f7;text-align:center;} .stat b{display:block;font-size:22px;color:#1a5276;}
    </style></head><body>`;
    if(logoSrc)html+=`<div style="display:flex;align-items:center;gap:14px;margin-bottom:16px;"><img src="${logoSrc}" width="55" style="border-radius:50%;background:#0d3249;"/><div><h1>${pName||""}</h1><p style="margin:2px 0;font-size:12px;color:#6c7a89;">${t.analytics} — ${label}</p><p style="margin:0;font-size:11px;color:#999;">${new Date().toLocaleDateString()}</p></div></div>`;
    if(expSec.summary){html+=`<div style="margin-bottom:16px;"><div class="stat"><b>${acts.length}</b>${t.totalActivities}</div><div class="stat"><b style="color:#27ae60">${acts.filter(a=>a.status==="Completed").length}</b>${t.completed}</div><div class="stat"><b style="color:#3498db">${acts.filter(a=>a.status==="Planned").length}</b>Planned</div><div class="stat"><b style="color:#1abc9c">${acts.filter(a=>a.status==="Approved").length}</b>Approved</div><div class="stat"><b style="color:#e74c3c">${overdueActs.length}</b>${t.overdue}</div><div class="stat"><b style="color:#7f8c8d">${acts.filter(a=>a.status==="Cancelled").length}</b>Cancelled</div><div class="stat"><b style="color:#95a5a6">${acts.filter(a=>a.status==="Postponed").length}</b>Postponed</div><div class="stat"><b style="color:#8e44ad">${acts.reduce((s,a)=>s+(a.attendees||0),0).toLocaleString()}</b>${t.totalAttendees}</div><div class="stat"><b style="color:#f39c12">${noRpt.length}</b>No Report</div></div>`;}
    if(expSec.monthly){html+=`<h2>📅 ${t.activitiesTimeline}</h2><table><tr><th>Month</th>`;MN.forEach(m=>{html+=`<th>${m}</th>`;});html+=`</tr><tr><td><b>Total</b></td>`;lineData.forEach(m=>{html+=`<td>${m.total}</td>`;});html+=`</tr><tr><td><b>${t.completed}</b></td>`;lineData.forEach(m=>{html+=`<td>${m.completed}</td>`;});html+=`</tr></table>`;}
    if(expSec.byType){html+=`<h2>📂 ${t.byType}</h2><table><tr><th>Type</th><th>Count</th></tr>`;typeData.forEach(x=>{html+=`<tr><td>${x.name}</td><td>${x.value}</td></tr>`;});html+=`</table>`;}
    if(expSec.collegeTable&&collegeStats.length>0){html+=`<h2>🏛️ College Performance</h2><table><tr><th>College</th><th>Total</th><th>Plans</th><th>Init.</th><th>Planned</th><th>Approved</th><th>Completed</th><th>Rate</th><th>Overdue</th><th>Cancelled</th><th>Postponed</th><th>No Report</th><th>Attendees</th><th>📰 Media</th></tr>`;collegeStats.forEach(c=>{const cls=c.rate>=80?"g":c.rate>=50?"y":"r";html+=`<tr><td>${c.name}</td><td>${c.total}</td><td>${c.plans}</td><td>${c.inits}</td><td>${c.planned}</td><td>${c.approved}</td><td>${c.completed}</td><td><span class="badge ${cls}">${c.rate}%</span></td><td style="color:${c.overdue>0?"#e74c3c":"#27ae60"};font-weight:${c.overdue>0?700:400}">${c.overdue}</td><td>${c.cancelled}</td><td>${c.postponed}</td><td style="color:${c.noReport>0?"#f39c12":"#27ae60"};font-weight:${c.noReport>0?700:400}">${c.noReport}</td><td>${c.attendees.toLocaleString()}</td><td>${c.mediaPub}</td></tr>`;});html+=`</table>`;}
    if(expSec.completed){const ca=acts.filter(a=>a.status==="Completed");if(ca.length>0){const rm={};(reports||[]).forEach(r=>{rm[r.activityTitle]=r;});html+=`<h2>✅ ${t.completed} (${ca.length})</h2><table><tr><th>#</th><th>Activity</th><th>Date</th><th>College</th><th>Presenter</th><th>Attendees</th><th>Report</th></tr>`;ca.forEach((a,i)=>{const rpt=rm[a.title]||{};const hasRpt=!!rpt.createdDate;html+=`<tr><td>${i+1}</td><td>${a.title}</td><td>${a.date||""}</td><td>${(a.college||"").replace("College of ","").replace("Department of ","")}</td><td>${a.presenter||rpt.presenter||"-"}</td><td>${a.attendees||rpt.totalParticipants||0}</td><td style="color:${hasRpt?"#27ae60":"#f39c12"};font-weight:700">${hasRpt?"✅ "+rpt.createdDate:"⚠️ No"}</td></tr>`;});html+=`</table>`;}}
    if(expSec.overdue&&overdueActs.length>0){html+=`<h2>⚠️ ${t.overdue} (${overdueActs.length})</h2><table><tr><th>#</th><th>Activity</th><th>Date</th><th>Days</th><th>College</th><th>Status</th></tr>`;overdueActs.sort((a,b)=>new Date(a.date)-new Date(b.date)).forEach((a,i)=>{html+=`<tr><td>${i+1}</td><td>${a.title}</td><td>${a.date}</td><td style="color:#e74c3c;font-weight:700">${Math.ceil((now-new Date(a.date))/86400000)}</td><td>${(a.college||"").replace("College of ","").replace("Department of ","")}</td><td>${a.status}</td></tr>`;});html+=`</table>`;}
    if(expSec.noReport&&noRpt.length>0){html+=`<h2>⚠️ Completed without report (${noRpt.length})</h2><table><tr><th>#</th><th>Activity</th><th>Date</th><th>College</th><th>Source</th></tr>`;noRpt.forEach((a,i)=>{html+=`<tr><td>${i+1}</td><td>${a.title}</td><td>${a.date||""}</td><td>${(a.college||"").replace("College of ","").replace("Department of ","")}</td><td>${a.source||""}</td></tr>`;});html+=`</table>`;}
    if(expSec.media){const rptM={};(reports||[]).forEach(r=>{rptM[r.activityTitle]=r;});const mItems=acts.filter(a=>{const rpt=rptM[a.title]||{};return a.mediaStatus==="Published"||a.mediaStatus==="Contacted"||rpt.mediaNotified||rpt.mediaLinks;});const published=acts.filter(a=>a.mediaStatus==="Published"||(rptM[a.title]?.mediaNotified&&rptM[a.title]?.mediaLinks)).length;if(mItems.length>0){html+=`<h2>📰 Media Coverage (${mItems.length} activities, ${published} published)</h2><table><tr><th>#</th><th>Activity</th><th>Date</th><th>College</th><th>Status</th><th>Channel</th><th>Links</th></tr>`;mItems.forEach((a,i)=>{const rpt=rptM[a.title]||{};html+=`<tr><td>${i+1}</td><td>${a.title}</td><td>${a.date||""}</td><td>${(a.college||"").replace("College of ","").replace("Department of ","")}</td><td>${a.mediaStatus||""}${rpt.mediaNotified?" ✅":""}</td><td>${a.mediaChannel||""}</td><td style="font-size:10px;">${rpt.mediaLinks||""}</td></tr>`;});html+=`</table>`;}else{html+=`<h2>📰 Media Coverage</h2><p>No media engagement recorded yet.</p>`;}}
    html+=`</body></html>`;
    showPrintPreview(html);
  };

  return(
    <div>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:16}}>
        <div><h2 style={{fontSize:18,fontWeight:700,margin:"0 0 3px"}}>{t.analytics}</h2><p style={{fontSize:12,color:"var(--color-text-secondary)",margin:0}}>{t.analyticsSubtitle}</p></div>
        <div style={{display:"flex",gap:8,alignItems:"center",flexWrap:"wrap"}}>
          <select value={fCollege} onChange={e=>setFCollege(e.target.value)} style={{...inp,width:"auto",minWidth:200,padding:"6px 10px",fontSize:12}}>
            <option value="All">🏛️ {t.all} — {allActs.length} {t.acts}</option>
            {(COLLEGES||[]).map(c=><option key={c} value={c}>{c.replace("College of ","").replace("Department of ","")} — {allActs.filter(a=>a.college===c).length}</option>)}
          </select>
          <button onClick={()=>setShowExportPanel(!showExportPanel)} style={{...btnP,padding:"5px 12px",fontSize:11,background:showExportPanel?"#7f8c8d":"#1a5276",display:"flex",alignItems:"center",gap:4}}>{showExportPanel?"✕ Close":"📤 Export"}</button>
        </div>
      </div>

      {/* EXPORT CUSTOMIZATION PANEL */}
      {showExportPanel&&<div style={{...crd,marginBottom:14,border:"2px solid #1a5276",padding:16}}>
        <h3 style={{fontSize:13,fontWeight:600,margin:"0 0 10px"}}>📤 Customize export — {fCollege==="All"?"All Colleges":fCollege.replace("College of ","").replace("Department of ","")}</h3>
        <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(170px,1fr))",gap:6,marginBottom:14}}>
          {[{k:"summary",l:"📊 Summary stats"},{k:"monthly",l:"📅 Monthly trend"},{k:"byType",l:"📂 By type"},{k:"byStatus",l:"🔄 By status"},{k:"collegeTable",l:"🏛️ College performance"},{k:"completed",l:"✅ Completed events"},{k:"overdue",l:"⚠️ Overdue list"},{k:"noReport",l:"⚠️ No report filed"},{k:"activityList",l:"📋 All activities"},{k:"media",l:"📰 Media coverage"}].map(s=>(
            <label key={s.k} style={{display:"flex",alignItems:"center",gap:6,fontSize:11,cursor:"pointer",padding:"5px 10px",borderRadius:6,background:expSec[s.k]?"#eaf2f8":"var(--color-background-secondary)",border:`1px solid ${expSec[s.k]?"#2980b9":"var(--color-border-tertiary)"}`}}>
              <input type="checkbox" checked={expSec[s.k]} onChange={()=>togSec(s.k)} style={{accentColor:"#1a5276"}}/>{s.l}
            </label>
          ))}
        </div>
        <div style={{display:"flex",gap:8}}>
          <button onClick={()=>{exportAnalyticsExcel();setShowExportPanel(false);}} style={{...btnP,padding:"8px 18px",fontSize:12,background:"#217346",display:"flex",alignItems:"center",gap:6}}>📊 Export Excel</button>
          <button onClick={()=>{printAnalytics();setShowExportPanel(false);}} style={{...btnP,padding:"8px 18px",fontSize:12,background:"#c0392b",display:"flex",alignItems:"center",gap:6}}>🖨️ Print / Save PDF</button>
          <button onClick={()=>{const all={};Object.keys(expSec).forEach(k=>{all[k]=true;});setExpSec(all);}} style={{...btnC,padding:"6px 12px",fontSize:11}}>Select all</button>
          <button onClick={()=>{const none={};Object.keys(expSec).forEach(k=>{none[k]=false;});setExpSec(none);}} style={{...btnC,padding:"6px 12px",fontSize:11}}>Deselect all</button>
        </div>
      </div>}
      {fCollege!=="All"&&<div style={{marginBottom:14,padding:"8px 14px",background:"#eaf2f8",borderRadius:8,fontSize:12,color:"#1a5276",fontWeight:500}}>🏛️ {fCollege} — {acts.length} {t.acts}, {acts.filter(a=>a.status==="Completed").length} {t.completed}, {acts.filter(a=>a.date&&new Date(a.date)<now&&a.status!=="Completed"&&a.status!=="Cancelled"&&a.status!=="Postponed").length} {t.overdue}</div>}

      {/* Summary stat cards */}
      {(()=>{
        const rptMap2={};(reports||[]).forEach(r=>{rptMap2[r.activityTitle]=r;});
        const mPub=acts.filter(a=>{const rpt=rptMap2[a.title]||{};return a.mediaStatus==="Published"||(rpt.mediaNotified&&rpt.mediaLinks);}).length;
        const planned2=acts.filter(a=>a.status==="Planned").length;
        const approved2=acts.filter(a=>a.status==="Approved").length;
        const inProg2=acts.filter(a=>a.status==="In Progress").length;
        const comp2=acts.filter(a=>a.status==="Completed").length;
        const canc2=acts.filter(a=>a.status==="Cancelled").length;
        const post2=acts.filter(a=>a.status==="Postponed").length;
        const ov2=acts.filter(a=>a.date&&new Date(a.date)<now&&a.status!=="Completed"&&a.status!=="Cancelled"&&a.status!=="Postponed").length;
        const att2=acts.reduce((s,a)=>s+(a.attendees||0),0);
        return <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(110px,1fr))",gap:8,marginBottom:14}}>
          <div style={{...crd,padding:"10px 12px",textAlign:"center"}}><div style={{fontSize:18,fontWeight:700}}>{acts.length}</div><div style={{fontSize:10,color:"var(--color-text-tertiary)"}}>{t.totalActivities}</div></div>
          <div style={{...crd,padding:"10px 12px",textAlign:"center"}}><div style={{fontSize:18,fontWeight:700,color:"#27ae60"}}>{comp2}</div><div style={{fontSize:10,color:"var(--color-text-tertiary)"}}>{t.completed} ({acts.length?Math.round(comp2/acts.length*100):0}%)</div></div>
          <div style={{...crd,padding:"10px 12px",textAlign:"center"}}><div style={{fontSize:18,fontWeight:700,color:"#3498db"}}>{planned2}</div><div style={{fontSize:10,color:"var(--color-text-tertiary)"}}>📝 Planned</div></div>
          <div style={{...crd,padding:"10px 12px",textAlign:"center"}}><div style={{fontSize:18,fontWeight:700,color:"#1abc9c"}}>{approved2}</div><div style={{fontSize:10,color:"var(--color-text-tertiary)"}}>✅ Approved</div></div>
          <div style={{...crd,padding:"10px 12px",textAlign:"center"}}><div style={{fontSize:18,fontWeight:700,color:"#f39c12"}}>{inProg2}</div><div style={{fontSize:10,color:"var(--color-text-tertiary)"}}>🔄 In Progress</div></div>
          <div style={{...crd,padding:"10px 12px",textAlign:"center"}}><div style={{fontSize:18,fontWeight:700,color:"#e74c3c"}}>{ov2}</div><div style={{fontSize:10,color:"var(--color-text-tertiary)"}}>⚠️ {t.overdue}</div></div>
          <div style={{...crd,padding:"10px 12px",textAlign:"center"}}><div style={{fontSize:18,fontWeight:700,color:"#95a5a6"}}>{post2}</div><div style={{fontSize:10,color:"var(--color-text-tertiary)"}}>⏸️ Postponed</div></div>
          <div style={{...crd,padding:"10px 12px",textAlign:"center"}}><div style={{fontSize:18,fontWeight:700,color:"#7f8c8d"}}>{canc2}</div><div style={{fontSize:10,color:"var(--color-text-tertiary)"}}>🚫 Cancelled</div></div>
          <div style={{...crd,padding:"10px 12px",textAlign:"center"}}><div style={{fontSize:18,fontWeight:700,color:"#8e44ad"}}>{att2.toLocaleString()}</div><div style={{fontSize:10,color:"var(--color-text-tertiary)"}}>👥 {t.totalAttendees}</div></div>
          <div style={{...crd,padding:"10px 12px",textAlign:"center"}}><div style={{fontSize:18,fontWeight:700,color:"#196f3d"}}>{mPub}</div><div style={{fontSize:10,color:"var(--color-text-tertiary)"}}>📰 Published ({acts.length?Math.round(mPub/acts.length*100):0}%)</div></div>
        </div>;
      })()}

      {/* ROW 1: Timeline + Type */}
      <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(300px,1fr))",gap:14,marginBottom:14}}>
        <div style={crd}><h3 style={{fontSize:13,fontWeight:600,margin:"0 0 10px"}}>{t.activitiesTimeline}</h3><div style={{height:200}}><ResponsiveContainer width="100%" height="100%"><LineChart data={lineData}><CartesianGrid strokeDasharray="3 3" opacity={0.3}/><XAxis dataKey="name" tick={{fontSize:9}}/><YAxis allowDecimals={false} tick={{fontSize:10}}/><Tooltip/><Legend wrapperStyle={{fontSize:11}}/><Line type="monotone" dataKey="total" stroke="#1a5276" strokeWidth={2} name={t.totalActivities}/><Line type="monotone" dataKey="completed" stroke="#27ae60" strokeWidth={2} name={t.completed}/></LineChart></ResponsiveContainer></div></div>
        <div style={crd}><h3 style={{fontSize:13,fontWeight:600,margin:"0 0 10px"}}>{t.byType}</h3><div style={{height:200}}><ResponsiveContainer width="100%" height="100%"><BarChart data={typeData} layout="vertical"><XAxis type="number" allowDecimals={false} tick={{fontSize:10}}/><YAxis dataKey="name" type="category" tick={{fontSize:9}} width={100}/><Tooltip/><Bar dataKey="value" fill="#c9a84c" radius={[0,3,3,0]}/></BarChart></ResponsiveContainer></div></div>
      </div>

      {/* ROW 2: Status + College total (plans vs initiatives) */}
      <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(300px,1fr))",gap:14,marginBottom:14}}>
        <div style={crd}><h3 style={{fontSize:13,fontWeight:600,margin:"0 0 10px"}}>{t.statusOverview}</h3><div style={{height:200}}><ResponsiveContainer width="100%" height="100%"><PieChart><Pie data={statusData} cx="50%" cy="50%" outerRadius={70} dataKey="value" label={({name,percent})=>`${name} ${Math.round(percent*100)}%`} labelLine={false} style={{fontSize:9}}>{statusData.map((e,i)=><Cell key={i} fill={sC(e.name)}/>)}</Pie><Tooltip/></PieChart></ResponsiveContainer></div></div>
        <div style={crd}><h3 style={{fontSize:13,fontWeight:600,margin:"0 0 10px"}}>🏛️ {t.college} — Plans vs Initiatives</h3><div style={{height:200}}><ResponsiveContainer width="100%" height="100%"><BarChart data={totalColData}><XAxis dataKey="name" tick={{fontSize:7}} angle={-25} textAnchor="end" height={45}/><YAxis allowDecimals={false} tick={{fontSize:10}}/><Tooltip/><Legend wrapperStyle={{fontSize:10}}/><Bar dataKey="plans" stackId="a" fill="#1a5276" name="Plans"/><Bar dataKey="inits" stackId="a" fill="#e67e22" name="Initiatives"/></BarChart></ResponsiveContainer></div></div>
      </div>

      {/* ROW 3: College comparison — Completed vs Overdue vs No Report */}
      <div style={crd}>
        <h3 style={{fontSize:13,fontWeight:600,margin:"0 0 10px"}}>📊 College comparison — Completed / Overdue / No report</h3>
        <div style={{height:220}}><ResponsiveContainer width="100%" height="100%"><BarChart data={compColData}><XAxis dataKey="name" tick={{fontSize:7}} angle={-25} textAnchor="end" height={45}/><YAxis allowDecimals={false} tick={{fontSize:10}}/><Tooltip/><Legend wrapperStyle={{fontSize:10}}/><Bar dataKey="completed" fill="#27ae60" name="✅ Completed" radius={[3,3,0,0]}/><Bar dataKey="overdue" fill="#e74c3c" name="⚠️ Overdue" radius={[3,3,0,0]}/><Bar dataKey="noReport" fill="#f39c12" name="⚠️ No report" radius={[3,3,0,0]}/><Bar dataKey="cancelled" fill="#7f8c8d" name="🚫 Cancelled" radius={[3,3,0,0]}/><Bar dataKey="postponed" fill="#95a5a6" name="⏸️ Postponed" radius={[3,3,0,0]}/></BarChart></ResponsiveContainer></div>
      </div>

      {/* ROW 4: Completion rate ranking + Attendance */}
      <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(300px,1fr))",gap:14,marginTop:14}}>
        <div style={crd}>
          <h3 style={{fontSize:13,fontWeight:600,margin:"0 0 10px"}}>🏆 Completion rate by college</h3>
          <div style={{height:200}}><ResponsiveContainer width="100%" height="100%"><BarChart data={rateColData.sort((a,b)=>b.rate-a.rate)} layout="vertical"><XAxis type="number" domain={[0,100]} tick={{fontSize:10}} unit="%"/><YAxis dataKey="name" type="category" tick={{fontSize:8}} width={90}/><Tooltip formatter={v=>v+"%"}/><Bar dataKey="rate" fill="#1a5276" radius={[0,3,3,0]}>{rateColData.sort((a,b)=>b.rate-a.rate).map((e,i)=><Cell key={i} fill={e.rate>=80?"#27ae60":e.rate>=50?"#f39c12":"#e74c3c"}/>)}</Bar></BarChart></ResponsiveContainer></div>
        </div>
        <div style={crd}>
          <h3 style={{fontSize:13,fontWeight:600,margin:"0 0 10px"}}>👥 Attendance by college</h3>
          <div style={{height:200}}><ResponsiveContainer width="100%" height="100%"><BarChart data={attColData.sort((a,b)=>b.attendees-a.attendees)} layout="vertical"><XAxis type="number" allowDecimals={false} tick={{fontSize:10}}/><YAxis dataKey="name" type="category" tick={{fontSize:8}} width={90}/><Tooltip/><Bar dataKey="attendees" fill="#8e44ad" radius={[0,3,3,0]}/></BarChart></ResponsiveContainer></div>
        </div>
      </div>

      {/* ROW 5: Detailed college table */}
      {collegeStats.length>0&&<div style={{...crd,marginTop:14}}>
        <h3 style={{fontSize:13,fontWeight:600,margin:"0 0 10px"}}>📋 College performance summary</h3>
        <div style={{overflowX:"auto"}}>
          <table style={{width:"100%",borderCollapse:"collapse",fontSize:11}}>
            <thead><tr style={{background:"#f4f6f7"}}>
              {["College","Total","Plans","Init.","Planned","Approved","Completed","Rate","Overdue","Cancelled","Postponed","No Report","Attendees","📰 Media"].map(h=><th key={h} style={{padding:"6px 8px",textAlign:isRtl?"right":"left",fontWeight:600,borderBottom:"2px solid #d5dbdb",whiteSpace:"nowrap"}}>{h}</th>)}
            </tr></thead>
            <tbody>{collegeStats.map((c,i)=>(
              <tr key={i} style={{borderBottom:"1px solid #eee"}}>
                <td style={{padding:"5px 8px",fontWeight:500}}>{c.name}</td>
                <td style={{padding:"5px 8px"}}>{c.total}</td>
                <td style={{padding:"5px 8px"}}>{c.plans}</td>
                <td style={{padding:"5px 8px",color:"#e67e22"}}>{c.inits}</td>
                <td style={{padding:"5px 8px",color:"#3498db"}}>{c.planned}</td>
                <td style={{padding:"5px 8px",color:"#1abc9c"}}>{c.approved}</td>
                <td style={{padding:"5px 8px",color:"#27ae60",fontWeight:600}}>{c.completed}</td>
                <td style={{padding:"5px 8px"}}><span style={{padding:"2px 6px",borderRadius:8,fontSize:10,fontWeight:600,background:c.rate>=80?"#d5f5e3":c.rate>=50?"#fef5e7":"#fdedec",color:c.rate>=80?"#196f3d":c.rate>=50?"#b7950b":"#c0392b"}}>{c.rate}%</span></td>
                <td style={{padding:"5px 8px",color:c.overdue>0?"#e74c3c":"#27ae60",fontWeight:c.overdue>0?700:400}}>{c.overdue}</td>
                <td style={{padding:"5px 8px",color:c.cancelled>0?"#7f8c8d":"#27ae60"}}>{c.cancelled}</td>
                <td style={{padding:"5px 8px",color:c.postponed>0?"#95a5a6":"#27ae60"}}>{c.postponed}</td>
                <td style={{padding:"5px 8px",color:c.noReport>0?"#f39c12":"#27ae60",fontWeight:c.noReport>0?700:400}}>{c.noReport}</td>
                <td style={{padding:"5px 8px"}}>{c.attendees.toLocaleString()}</td>
                <td style={{padding:"5px 8px",color:c.mediaPub>0?"#196f3d":"#888"}}>{c.mediaPub}</td>
              </tr>
            ))}</tbody>
          </table>
        </div>
      </div>}

      {/* ROW 6: MEDIA COVERAGE ANALYTICS */}
      {(()=>{
        // Compute media stats from activities + reports
        const rptMap={};(reports||[]).forEach(r=>{rptMap[r.activityTitle]=r;});
        const mediaStats={published:0,contacted:0,notContacted:0,declined:0,withLinks:0,channels:{}};
        const mediaByCollege={};
        const mediaItems=[];
        acts.forEach(a=>{
          const rpt=rptMap[a.title]||{};
          const status=a.mediaStatus||rpt.mediaStatus||"Not contacted";
          const notified=rpt.mediaNotified||false;
          const links=a.mediaLinks||rpt.mediaLinks||"";
          const channel=a.mediaChannel||rpt.mediaChannel||"";
          const isPublished=status==="Published"||(notified&&links);
          const isContacted=status==="Contacted"||status==="Coverage scheduled"||notified;

          if(isPublished)mediaStats.published++;
          else if(status==="Declined")mediaStats.declined++;
          else if(isContacted)mediaStats.contacted++;
          else mediaStats.notContacted++;
          if(links)mediaStats.withLinks++;
          if(channel){mediaStats.channels[channel]=(mediaStats.channels[channel]||0)+1;}
          if(isPublished||isContacted||links){
            mediaItems.push({title:a.title,date:a.date,college:(a.college||"").replace("College of ","").replace("Department of ",""),status,channel,links,notified});
          }
          // By college
          const col=(a.college||"Unknown").replace("College of ","").replace("Department of ","");
          if(!mediaByCollege[col])mediaByCollege[col]={published:0,contacted:0,none:0};
          if(isPublished)mediaByCollege[col].published++;
          else if(isContacted)mediaByCollege[col].contacted++;
          else mediaByCollege[col].none++;
        });
        const totalMedia=mediaStats.published+mediaStats.contacted+mediaStats.declined;
        const mediaPieData=[
          {name:"Published",value:mediaStats.published},
          {name:"Contacted",value:mediaStats.contacted},
          {name:"Declined",value:mediaStats.declined},
          {name:"Not contacted",value:mediaStats.notContacted}
        ].filter(d=>d.value>0);
        const mediaColData=Object.entries(mediaByCollege).filter(([,v])=>v.published+v.contacted>0).map(([k,v])=>({name:k,...v}));
        const channelData=Object.entries(mediaStats.channels).map(([k,v])=>({name:k,value:v})).sort((a,b)=>b.value-a.value);
        const MEDIA_COLORS=["#27ae60","#3498db","#e74c3c","#95a5a6"];

        return(
          <div style={{...crd,marginTop:14}}>
            <h3 style={{fontSize:14,fontWeight:700,margin:"0 0 14px",color:"#1a5276"}}>📰 Media Coverage Analytics</h3>
            {/* Summary badges */}
            <div style={{display:"flex",gap:8,flexWrap:"wrap",marginBottom:14}}>
              <span style={{fontSize:11,padding:"4px 12px",borderRadius:8,background:"#d5f5e3",color:"#196f3d",fontWeight:600}}>📰 {mediaStats.published} Published</span>
              <span style={{fontSize:11,padding:"4px 12px",borderRadius:8,background:"#eaf2f8",color:"#2471a3",fontWeight:600}}>📞 {mediaStats.contacted} Contacted</span>
              <span style={{fontSize:11,padding:"4px 12px",borderRadius:8,background:"#fdedec",color:"#c0392b",fontWeight:600}}>❌ {mediaStats.declined} Declined</span>
              <span style={{fontSize:11,padding:"4px 12px",borderRadius:8,background:"#f5f5f5",color:"#888",fontWeight:600}}>⬜ {mediaStats.notContacted} Not contacted</span>
              {mediaStats.withLinks>0&&<span style={{fontSize:11,padding:"4px 12px",borderRadius:8,background:"#f0e6f6",color:"#6c3483",fontWeight:600}}>🔗 {mediaStats.withLinks} with links</span>}
              {acts.length>0&&<span style={{fontSize:11,padding:"4px 12px",borderRadius:8,background:"#fef5e7",color:"#b7950b",fontWeight:600}}>📊 {acts.length>0?Math.round((mediaStats.published/acts.length)*100):0}% coverage rate</span>}
            </div>
            {/* Charts row */}
            <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(250px,1fr))",gap:14,marginBottom:14}}>
              {mediaPieData.length>0&&<div>
                <div style={{fontSize:12,fontWeight:600,marginBottom:6}}>Coverage status</div>
                <div style={{height:180}}><ResponsiveContainer width="100%" height="100%"><PieChart><Pie data={mediaPieData} cx="50%" cy="50%" outerRadius={65} dataKey="value" label={({name,percent})=>`${name} ${Math.round(percent*100)}%`} labelLine={false} style={{fontSize:9}}>{mediaPieData.map((e,i)=><Cell key={i} fill={MEDIA_COLORS[i%MEDIA_COLORS.length]}/>)}</Pie><Tooltip/></PieChart></ResponsiveContainer></div>
              </div>}
              {mediaColData.length>0&&<div>
                <div style={{fontSize:12,fontWeight:600,marginBottom:6}}>Media by college</div>
                <div style={{height:180}}><ResponsiveContainer width="100%" height="100%"><BarChart data={mediaColData}><XAxis dataKey="name" tick={{fontSize:7}} angle={-25} textAnchor="end" height={45}/><YAxis allowDecimals={false} tick={{fontSize:10}}/><Tooltip/><Legend wrapperStyle={{fontSize:10}}/><Bar dataKey="published" stackId="a" fill="#27ae60" name="Published"/><Bar dataKey="contacted" stackId="a" fill="#3498db" name="Contacted"/></BarChart></ResponsiveContainer></div>
              </div>}
              {channelData.length>0&&<div>
                <div style={{fontSize:12,fontWeight:600,marginBottom:6}}>Channels / Platforms</div>
                <div style={{height:180}}><ResponsiveContainer width="100%" height="100%"><BarChart data={channelData} layout="vertical"><XAxis type="number" allowDecimals={false} tick={{fontSize:10}}/><YAxis dataKey="name" type="category" tick={{fontSize:9}} width={80}/><Tooltip/><Bar dataKey="value" fill="#8e44ad" radius={[0,3,3,0]}/></BarChart></ResponsiveContainer></div>
              </div>}
            </div>
            {/* Media items list */}
            {mediaItems.length>0&&<div>
              <div style={{fontSize:12,fontWeight:600,marginBottom:6}}>Activities with media engagement ({mediaItems.length})</div>
              <div style={{maxHeight:200,overflowY:"auto",border:"1px solid var(--color-border-tertiary)",borderRadius:8}}>
                {mediaItems.map((m,i)=>(
                  <div key={i} style={{display:"flex",alignItems:"center",gap:8,padding:"5px 10px",borderBottom:i<mediaItems.length-1?"0.5px solid var(--color-border-tertiary)":"none",fontSize:11}}>
                    <span style={{fontSize:9,padding:"2px 6px",borderRadius:6,background:m.status==="Published"?"#d5f5e3":m.status==="Contacted"?"#eaf2f8":"#f5f5f5",color:m.status==="Published"?"#196f3d":m.status==="Contacted"?"#2471a3":"#888",fontWeight:600,flexShrink:0}}>{m.status==="Published"?"📰":m.notified?"📞":"📝"}</span>
                    <span style={{flex:1,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{m.title}</span>
                    <span style={{fontSize:9,color:"var(--color-text-tertiary)",flexShrink:0}}>{m.college}</span>
                    <span style={{fontSize:9,color:"var(--color-text-tertiary)",flexShrink:0}}>{m.date||""}</span>
                    {m.channel&&<span style={{fontSize:9,padding:"1px 5px",borderRadius:4,background:"#f0e6f6",color:"#6c3483",flexShrink:0}}>{m.channel}</span>}
                  </div>
                ))}
              </div>
            </div>}
            {totalMedia===0&&mediaStats.notContacted>0&&<div style={{padding:"12px 16px",background:"#fef9e7",borderRadius:8,fontSize:12,color:"#7d6608"}}>ℹ️ No media engagement yet. {mediaStats.notContacted} activities have not been contacted for media coverage.</div>}
          </div>
        );
      })()}
    </div>
  );
}

// ─── PLANS ────────────────────────────────────────────────────────────
function PlansPage({t,plans,savePlans,user,canEdit,isAdmin,onSelect,COLLEGES}){
  const [sf,setSf]=useState(false);
  const vis=user.role==="college"?plans.filter(p=>p.college===user.college):plans;
  const add=d=>{savePlans([...plans,{id:gid(),...d,activities:[],submittedDate:new Date().toISOString().split("T")[0]}]);setSf(false);};
  return(
    <div>
      <div style={{display:"flex",justifyContent:"space-between",marginBottom:16}}>
        <div><h2 style={{fontSize:18,fontWeight:700,margin:0}}>{t.yearlyPlans}</h2><p style={{fontSize:12,color:"var(--color-text-secondary)",margin:"3px 0 0"}}>{t.planSubtitle}</p></div>
        {canEdit&&<button onClick={()=>setSf(true)} style={{...btnP,display:"flex",alignItems:"center",gap:5}}><Ic n="plus" s={15} c="#fff"/>{t.newPlan}</button>}
      </div>
      {sf&&<FrmCard t={t} title={t.newPlan} onCancel={()=>setSf(false)} onSave={add} fields={[{k:"college",l:t.college,type:user.role==="college"?"fixed":"select",v:user.role==="college"?user.college:"",opts:COLLEGES},{k:"year",l:t.year,type:"select",v:"2025-2026",opts:["2024-2025","2025-2026","2026-2027","2027-2028"]},{k:"status",l:t.status,type:"fixed",v:"Pending Review"}]}/>}
      {vis.map(p=>(
        <div key={p.id} onClick={()=>onSelect(p)} style={{...crd,marginBottom:10,cursor:"pointer",borderLeft:`3px solid ${sC(p.status)}`}}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}>
            <div><div style={{fontSize:14,fontWeight:600}}>{p.college}</div><div style={{fontSize:11,color:"var(--color-text-secondary)"}}>{t.submittedDate}: {p.submittedDate} · {t.year}: {p.year}</div></div>
            <div style={{display:"flex",alignItems:"center",gap:8}}>
              <span style={{fontSize:10,padding:"3px 10px",borderRadius:20,background:`${sC(p.status)}18`,color:sC(p.status),fontWeight:600}}>{p.status}</span>
              {isAdmin&&p.status==="Pending Review"&&<button onClick={e=>{e.stopPropagation();savePlans(plans.map(x=>x.id===p.id?{...x,status:"Approved"}:x));}} style={{...btnP,padding:"4px 10px",fontSize:11,background:"#27ae60"}}>{t.approve}</button>}
              <span style={{fontSize:12,fontWeight:700,color:"#1a5276"}}>{p.activities.length} {t.acts}</span>
            </div>
          </div>
        </div>
      ))}
      {vis.length===0&&<div style={{textAlign:"center",padding:50,color:"var(--color-text-tertiary)"}}>{t.noPlans}</div>}
    </div>
  );
}

function PlanDetail({t,plan,plans,savePlans,canEdit,isAdmin,user,onBack,trainers,logoSrc,pName,pSub}){
  const [sa,setSa]=useState(false);
  const [bulkExtracting,setBulkExtracting]=useState(false);
  const [bulkResults,setBulkResults]=useState(null);
  const [bulkMsg,setBulkMsg]=useState("");
  const cm=isAdmin||(user.role==="college"&&user.college===plan.college);
  const add=d=>{const a={id:gid(),...d,attendees:0,mediaNotified:false,mediaLinks:"",participants:""};savePlans(plans.map(p=>p.id===plan.id?{...p,activities:[...p.activities,a]}:p));setSa(false);};
  const addBulk=acts=>{const newActs=acts.map(a=>({id:gid(),title:a.title||"",type:ACT_TYPES.includes(a.type)?a.type:"Lecture",date:a.date||"",presenter:a.presenter||"",delivery:DELIVERY.includes(a.delivery)?a.delivery:"In-Person",target:TARGETS.includes(a.target)?a.target:"All",venue:a.venue||"",duration:a.duration||"",objectives:a.objectives||"",notes:a.notes||"",status:"Planned",source:"College Plan",department:a.department||"",mediaStatus:"Not contacted",mediaContactedDate:"",mediaPublishedDate:"",mediaChannel:"",mediaCoverageNotes:"",attendees:0,mediaNotified:false,mediaLinks:"",participants:"",customFields:a.customFields||[]}));savePlans(plans.map(p=>p.id===plan.id?{...p,activities:[...p.activities,...newActs]}:p));setBulkResults(null);setBulkMsg("");};
  const upd=(aid,u)=>savePlans(plans.map(p=>p.id===plan.id?{...p,activities:p.activities.map(a=>a.id===aid?{...a,...u}:a)}:p));
  const del=aid=>savePlans(plans.map(p=>p.id===plan.id?{...p,activities:p.activities.filter(a=>a.id!==aid)}:p));
  const handleBulkFile=async file=>{
    if(!file)return;setBulkExtracting(true);setBulkMsg("");setBulkResults(null);
    try{
      if(file.size>4*1024*1024){setBulkMsg("⚠️ File too large (max 4MB).");setBulkExtracting(false);return;}
      const mtype=file.type||"application/octet-stream";
      const isSupported=mtype.startsWith("image/")||mtype==="application/pdf"||mtype.startsWith("text/");
      if(!isSupported){setBulkMsg("⚠️ Use PDF, image, or camera. .docx must be converted to PDF first.");setBulkExtracting(false);return;}
      const b64=await new Promise((res,rej)=>{const r=new FileReader();r.onload=()=>res(r.result.split(",")[1]);r.onerror=()=>rej(new Error("File read failed"));r.readAsDataURL(file);});
      const acts=await extractBulkActivities(b64,mtype,t.dir==="rtl"?"ar":"en");
      if(acts&&acts.length>0){setBulkResults(acts);setBulkMsg(`✅ ${acts.length} ${t.activitiesFound}`);}
      else{setBulkMsg("❌ "+t.extractionError+" (Check browser console for details)");}
    }catch(e){console.error("Bulk upload error:",e);setBulkMsg("❌ Error: "+e.message);}
    setBulkExtracting(false);
  };
  const removeBulkItem=i=>setBulkResults(bulkResults.filter((_,idx)=>idx!==i));

  // Get current plan fresh from plans array
  const currentPlan = plans.find(p=>p.id===plan.id) || plan;

  return(
    <div>
      <button onClick={onBack} style={{background:"none",border:"none",cursor:"pointer",fontSize:12,color:"var(--color-text-secondary)",padding:0,marginBottom:12}}>← {t.back}</button>
      <div style={{display:"flex",justifyContent:"space-between",marginBottom:16,flexWrap:"wrap",gap:8}}>
        <div><h2 style={{fontSize:18,fontWeight:700,margin:0}}>{currentPlan.college}</h2><p style={{fontSize:12,color:"var(--color-text-secondary)",margin:"3px 0 0"}}>{t.year}: {currentPlan.year} · <span style={{color:sC(currentPlan.status)}}>{currentPlan.status}</span> · {currentPlan.activities.length} {t.acts}</p></div>
        {cm&&<div style={{display:"flex",gap:8}}>
          <button onClick={()=>setSa(true)} style={{...btnP,display:"flex",alignItems:"center",gap:5}}><Ic n="plus" s={15} c="#fff"/>{t.newActivity}</button>
        </div>}
      </div>

      {/* BULK IMPORT */}
      {cm&&<div style={{background:"var(--color-background-secondary)",borderRadius:10,padding:16,marginBottom:14,border:"1px dashed var(--color-border-secondary)"}}>
        <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:8}}>
          <span style={{fontSize:20}}>📋</span>
          <div><div style={{fontSize:13,fontWeight:600}}>{t.bulkImport}</div><div style={{fontSize:11,color:"var(--color-text-tertiary)"}}>{t.bulkImportHint}</div></div>
        </div>
        <div style={{display:"flex",gap:10,alignItems:"center",flexWrap:"wrap"}}>
          <label style={{...btnP,padding:"7px 16px",fontSize:12,display:"inline-flex",alignItems:"center",gap:5,cursor:"pointer",opacity:bulkExtracting?0.5:1}}>
            <Ic n="plan" s={14} c="#fff"/>{t.fileUpload}
            <input type="file" accept="image/*,.pdf,.txt" onChange={e=>handleBulkFile(e.target.files[0])} style={{display:"none"}} disabled={bulkExtracting}/>
          </label>
          <label style={{...btnP,padding:"7px 16px",fontSize:12,display:"inline-flex",alignItems:"center",gap:5,cursor:"pointer",background:"#27ae60",opacity:bulkExtracting?0.5:1}}>
            📷 {t.camera}
            <input type="file" accept="image/*" capture="environment" onChange={e=>handleBulkFile(e.target.files[0])} style={{display:"none"}} disabled={bulkExtracting}/>
          </label>
          {bulkExtracting&&<span style={{fontSize:12,color:"#e67e22",fontWeight:600}}>⏳ {t.extractingBulk}</span>}
        </div>
        {bulkMsg&&<div style={{marginTop:8,fontSize:12,color:bulkResults?"#27ae60":"#e74c3c",fontWeight:500}}>{bulkMsg}</div>}

        {/* Preview extracted activities */}
        {bulkResults&&bulkResults.length>0&&<div style={{marginTop:12}}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:8}}>
            <span style={{fontSize:12,fontWeight:600}}>{bulkResults.length} {t.activitiesFound}:</span>
            <button onClick={()=>addBulk(bulkResults)} style={{...btnP,padding:"6px 16px",fontSize:12}}>{t.importAll} ({bulkResults.length})</button>
          </div>
          {bulkResults.map((a,i)=>(
            <div key={i} style={{background:"#ffffff",borderRadius:8,padding:"10px 14px",marginBottom:6,border:"0.5px solid var(--color-border-tertiary)",display:"flex",justifyContent:"space-between",alignItems:"center"}}>
              <div style={{flex:1,minWidth:0}}>
                <div style={{fontSize:12,fontWeight:600}}>{a.title||`Activity ${i+1}`}</div>
                <div style={{fontSize:11,color:"var(--color-text-secondary)"}}>
                  {a.type&&<span>{a.type} · </span>}{a.date&&<span>{a.date} · </span>}{a.presenter&&<span>{a.presenter} · </span>}{a.department&&<span>🏢 {a.department}</span>}
                </div>
              </div>
              <button onClick={()=>removeBulkItem(i)} style={btnI}><Ic n="trash" s={13} c="#e74c3c"/></button>
            </div>
          ))}
        </div>}
      </div>}

      {sa&&<ActForm t={t} onSave={add} onCancel={()=>setSa(false)} trainers={trainers} source="College Plan"/>}
      {currentPlan.activities.map(a=><ActCard key={a.id} t={t} a={{...a,college:currentPlan.college}} canEdit={cm} onUpdate={u=>upd(a.id,u)} onDelete={()=>del(a.id)} trainers={trainers} logoSrc={logoSrc} pName={pName} pSub={pSub}/>)}
      {currentPlan.activities.length===0&&!bulkResults&&<div style={{textAlign:"center",padding:40,color:"var(--color-text-tertiary)",fontSize:12}}>{t.noActivities}</div>}
    </div>
  );
}

// ─── ACTIVITY FORM ────────────────────────────────────────────────────
function ActForm({t,onSave,onCancel,trainers,source,initial,COLLEGES}){
  const knownKeys=["title","type","date","presenter","delivery","target","venue","duration","objectives","notes","status","source","department","college","mediaStatus","mediaContactedDate","mediaPublishedDate","mediaChannel","mediaCoverageNotes"];
  const [d,sD]=useState({title:initial?.title||"",type:initial?.type||"Lecture",date:initial?.date||"",presenter:initial?.presenter||"",delivery:initial?.delivery||"In-Person",target:initial?.target||"Students",venue:initial?.venue||"",duration:initial?.duration||"",objectives:initial?.objectives||"",notes:initial?.notes||"",status:initial?.status||"Planned",source:source||initial?.source||"College Plan",department:initial?.department||"",college:initial?.college||"",mediaStatus:initial?.mediaStatus||"Not contacted",mediaContactedDate:initial?.mediaContactedDate||"",mediaPublishedDate:initial?.mediaPublishedDate||"",mediaChannel:initial?.mediaChannel||"",mediaCoverageNotes:initial?.mediaCoverageNotes||""});
  const [customFields,setCF]=useState(initial?.customFields||[]);
  const [extracting,setExtracting]=useState(false);
  const [extractMsg,setExtractMsg]=useState("");
  const s=(k,v)=>sD({...d,[k]:v});
  const addCF=()=>setCF([...customFields,{id:gid(),name:"",value:""}]);
  const updCF=(id,field,val)=>setCF(customFields.map(f=>f.id===id?{...f,[field]:val}:f));
  const delCF=id=>setCF(customFields.filter(f=>f.id!==id));

  const handleFileUpload=async(file)=>{
    if(!file)return;
    setExtracting(true);setExtractMsg("");
    try{
      // Check file size (limit to ~4MB for API)
      if(file.size>4*1024*1024){setExtractMsg("⚠️ File too large (max 4MB). Try a smaller file or image.");setExtracting(false);return;}
      const b64=await new Promise((res,rej)=>{const r=new FileReader();r.onload=()=>res(r.result.split(",")[1]);r.onerror=()=>rej(new Error("File read failed"));r.readAsDataURL(file);});
      // For .docx and other binary formats, treat as image by asking user to use PDF/image instead
      const mtype=file.type||"application/octet-stream";
      const isSupported=mtype.startsWith("image/")||mtype==="application/pdf"||mtype.startsWith("text/");
      if(!isSupported){setExtractMsg("⚠️ .docx/.doc files must be converted to PDF or photographed first. Use PDF, image, or camera.");setExtracting(false);return;}
      const result=await extractFromFile(b64,mtype,t.dir==="rtl"?"ar":"en");
      if(result){
        const newD={...d};
        knownKeys.forEach(k=>{if(result[k]&&typeof result[k]==="string"&&result[k].trim())newD[k]=result[k].trim();});
        if(result.type){const mt=ACT_TYPES.find(x=>x.toLowerCase()===result.type.toLowerCase());if(mt)newD.type=mt;}
        if(result.source){const ms=SOURCES.find(x=>x.toLowerCase()===result.source.toLowerCase());if(ms)newD.source=ms;}
        if(result.delivery){const md=DELIVERY.find(x=>x.toLowerCase()===result.delivery.toLowerCase());if(md)newD.delivery=md;}
        if(result.target){const mt2=TARGETS.find(x=>x.toLowerCase()===result.target.toLowerCase());if(mt2)newD.target=mt2;}
        sD(newD);
        if(result.customFields&&Array.isArray(result.customFields)&&result.customFields.length>0){
          setCF(prev=>[...prev,...result.customFields.map(cf=>({id:gid(),name:cf.name||"",value:cf.value||""}))]);
        }
        setExtractMsg(t.extractionDone);
      }else{setExtractMsg("❌ "+t.extractionError+" (Check browser console for details)");}
    }catch(e){console.error("Upload handler error:",e);setExtractMsg("❌ Error: "+e.message);}
    setExtracting(false);
  };

  const doSave=()=>{if(!d.title&&!d.date)return;const out={...d,customFields:customFields.filter(f=>f.name.trim())};onSave(out);};

  return(
    <div style={{...crd,border:"1.5px solid #c9a84c",marginBottom:16}}>
      <h3 style={{fontSize:14,fontWeight:600,margin:"0 0 8px"}}>{initial?t.edit:t.newActivity}</h3>

      {/* FILE UPLOAD / AI EXTRACTION */}
      {!initial&&<div style={{background:"var(--color-background-secondary)",borderRadius:10,padding:16,marginBottom:14,border:"1px dashed var(--color-border-secondary)"}}>
        <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:8}}>
          <span style={{fontSize:20}}>📄</span>
          <div><div style={{fontSize:13,fontWeight:600}}>{t.extractFromFile}</div><div style={{fontSize:11,color:"var(--color-text-tertiary)"}}>{t.uploadHint}</div></div>
        </div>
        <div style={{display:"flex",gap:10,alignItems:"center",flexWrap:"wrap"}}>
          <label style={{...btnP,padding:"7px 16px",fontSize:12,display:"inline-flex",alignItems:"center",gap:5,cursor:"pointer",opacity:extracting?0.5:1}}>
            <Ic n="plan" s={14} c="#fff"/>{t.fileUpload}
            <input type="file" accept="image/*,.pdf,.txt" onChange={e=>handleFileUpload(e.target.files[0])} style={{display:"none"}} disabled={extracting}/>
          </label>
          <label style={{...btnP,padding:"7px 16px",fontSize:12,display:"inline-flex",alignItems:"center",gap:5,cursor:"pointer",background:"#27ae60",opacity:extracting?0.5:1}}>
            📷 {t.camera}
            <input type="file" accept="image/*" capture="environment" onChange={e=>handleFileUpload(e.target.files[0])} style={{display:"none"}} disabled={extracting}/>
          </label>
          {extracting&&<span style={{fontSize:12,color:"#e67e22",fontWeight:600}}>⏳ {t.extracting}</span>}
        </div>
        {extractMsg&&<div style={{marginTop:8,fontSize:12,color:extractMsg===t.extractionDone?"#27ae60":"#e74c3c",fontWeight:500}}>{extractMsg}</div>}
        <div style={{fontSize:11,color:"var(--color-text-tertiary)",marginTop:6}}>{t.orFillManually}</div>
      </div>}

      <p style={{fontSize:10,color:"var(--color-text-tertiary)",margin:"0 0 10px"}}>{t.guideline10days}</p>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>
        <div style={{gridColumn:"1/-1"}}><label style={lbl}>{t.title} (عنوان النشاط)</label><input value={d.title} onChange={e=>s("title",e.target.value)} style={inp}/></div>
        <div><label style={lbl}>{t.type}</label><select value={d.type} onChange={e=>s("type",e.target.value)} style={inp}>{ACT_TYPES.map(x=><option key={x}>{x}</option>)}</select></div>
        <div><label style={lbl}>{t.date}</label><input type="date" value={d.date} onChange={e=>s("date",e.target.value)} style={inp}/></div>
        <div><label style={lbl}>{t.presenter}</label><select value={d.presenter} onChange={e=>s("presenter",e.target.value)} style={inp}><option value="">--</option>{trainers.map(x=><option key={x.id} value={x.name}>{x.name}</option>)}</select><input value={d.presenter} onChange={e=>s("presenter",e.target.value)} placeholder="..." style={{...inp,marginTop:3}}/></div>
        <div><label style={lbl}>{t.delivery}</label><select value={d.delivery} onChange={e=>s("delivery",e.target.value)} style={inp}>{DELIVERY.map(x=><option key={x}>{x}</option>)}</select></div>
        <div><label style={lbl}>{t.target}</label><select value={d.target} onChange={e=>s("target",e.target.value)} style={inp}>{TARGETS.map(x=><option key={x}>{x}</option>)}</select></div>
        <div><label style={lbl}>{t.venue}</label><input value={d.venue} onChange={e=>s("venue",e.target.value)} style={inp}/></div>
        <div><label style={lbl}>{t.duration}</label><input value={d.duration} onChange={e=>s("duration",e.target.value)} style={inp}/></div>
        <div><label style={lbl}>{t.department}</label><input value={d.department} onChange={e=>s("department",e.target.value)} placeholder={t.noDept} style={inp}/></div>
        {COLLEGES&&COLLEGES.length>0&&<div><label style={lbl}>{t.college}</label><select value={d.college} onChange={e=>s("college",e.target.value)} style={inp}><option value="">--</option>{COLLEGES.map(c=><option key={c} value={c}>{c}</option>)}</select></div>}
        <div><label style={lbl}>{t.status}</label><select value={d.status} onChange={e=>s("status",e.target.value)} style={inp}>{STATUSES.map(x=><option key={x}>{x}</option>)}</select></div>
        {!source&&<div><label style={lbl}>{t.source}</label><select value={d.source} onChange={e=>s("source",e.target.value)} style={inp}>{SOURCES.map(x=><option key={x}>{x}</option>)}</select></div>}
        <div style={{gridColumn:"1/-1"}}><label style={lbl}>{t.objectives}</label><textarea value={d.objectives} onChange={e=>s("objectives",e.target.value)} rows={2} style={{...inp,resize:"vertical"}}/></div>
        <div style={{gridColumn:"1/-1"}}><label style={lbl}>{t.notes}</label><textarea value={d.notes} onChange={e=>s("notes",e.target.value)} rows={2} style={{...inp,resize:"vertical"}}/></div>
        <div style={{gridColumn:"1/-1",borderTop:"1px dashed var(--color-border-tertiary)",paddingTop:10,marginTop:4}}><span style={{fontSize:12,fontWeight:600,color:"#e67e22"}}>📺 {t.mediaCoverage}</span></div>
        <div><label style={lbl}>{t.mediaStatus}</label><select value={d.mediaStatus} onChange={e=>s("mediaStatus",e.target.value)} style={inp}><option value="Not contacted">{t.notContacted}</option><option value="Contacted">{t.contacted}</option><option value="Coverage scheduled">{t.coverageScheduled}</option><option value="Published">{t.published}</option><option value="Declined">{t.declined}</option></select></div>
        <div><label style={lbl}>{t.mediaChannel}</label><input value={d.mediaChannel} onChange={e=>s("mediaChannel",e.target.value)} placeholder="Facebook, YouTube, Website..." style={inp}/></div>
        <div><label style={lbl}>{t.mediaContactedDate}</label><input type="date" value={d.mediaContactedDate} onChange={e=>s("mediaContactedDate",e.target.value)} style={inp}/></div>
        <div><label style={lbl}>{t.mediaPublishedDate}</label><input type="date" value={d.mediaPublishedDate} onChange={e=>s("mediaPublishedDate",e.target.value)} style={inp}/></div>
        <div style={{gridColumn:"1/-1"}}><label style={lbl}>{t.mediaCoverageNotes}</label><textarea value={d.mediaCoverageNotes} onChange={e=>s("mediaCoverageNotes",e.target.value)} rows={2} placeholder={t.mediaCoverageNotes+"..."} style={{...inp,resize:"vertical"}}/></div>

        {/* CUSTOM / EXTRACTED FIELDS */}
        {customFields.length>0&&<div style={{gridColumn:"1/-1",borderTop:"1px dashed var(--color-border-tertiary)",paddingTop:10,marginTop:4}}><span style={{fontSize:12,fontWeight:600,color:"#8e44ad"}}>📋 {t.extractedFields}</span></div>}
        {customFields.map(cf=>(
          <div key={cf.id} style={{gridColumn:"1/-1",display:"flex",gap:8,alignItems:"flex-start"}}>
            <div style={{flex:1}}><label style={lbl}>{t.fieldName}</label><input value={cf.name} onChange={e=>updCF(cf.id,"name",e.target.value)} placeholder={t.fieldName} style={inp}/></div>
            <div style={{flex:2}}><label style={lbl}>{t.fieldValue}</label><input value={cf.value} onChange={e=>updCF(cf.id,"value",e.target.value)} placeholder={t.fieldValue} style={inp}/></div>
            <button onClick={()=>delCF(cf.id)} style={{...btnI,marginTop:18,color:"#e74c3c"}} title={t.removeField}><Ic n="trash" s={14} c="#e74c3c"/></button>
          </div>
        ))}
        <div style={{gridColumn:"1/-1"}}><button onClick={addCF} style={{...btnC,fontSize:11,padding:"5px 12px",display:"inline-flex",alignItems:"center",gap:4}}><Ic n="plus" s={13}/>{t.addCustomField}</button></div>
      </div>
      <div style={{display:"flex",gap:8,marginTop:12,justifyContent:"flex-end"}}><button onClick={onCancel} style={btnC}>{t.cancel}</button><button onClick={doSave} disabled={!d.title||!d.date} style={{...btnP,opacity:d.title&&d.date?1:0.5}}>{t.save}</button></div>
    </div>
  );
}

function ActCard({t,a,canEdit,onUpdate,onDelete,trainers=[],logoSrc,pName,pSub,COLLEGES}){
  const [ed,sEd]=useState(false);
  const [showExport,setShowExport]=useState(false);
  if(ed&&canEdit)return <ActForm t={t} initial={a} onSave={d=>{onUpdate(d);sEd(false);}} onCancel={()=>sEd(false)} trainers={trainers} COLLEGES={COLLEGES}/>;
  return(
    <div style={{...crd,marginBottom:8,borderLeft:`3px solid ${sC(a.status)}`,padding:"12px 16px"}}>
      <div style={{display:"flex",justifyContent:"space-between",gap:10}}>
        <div style={{flex:1,minWidth:0}}>
          <div style={{display:"flex",gap:5,flexWrap:"wrap",marginBottom:3,alignItems:"center"}}>
            <span style={{fontSize:13,fontWeight:600}}>{a.title}</span>
            <span style={{fontSize:9,padding:"2px 7px",borderRadius:10,background:"var(--color-background-secondary)",color:"var(--color-text-secondary)"}}>{a.type}</span>
            <span style={{fontSize:9,padding:"2px 7px",borderRadius:10,background:`${sC(a.status)}18`,color:sC(a.status),fontWeight:600}}>{a.status}</span>
            {a.delivery&&<span style={{fontSize:9,padding:"2px 7px",borderRadius:10,background:"#f0f4f7",color:"#5d6d7e"}}>{a.delivery}</span>}
          </div>
          <div style={{display:"flex",gap:12,fontSize:11,color:"var(--color-text-secondary)",flexWrap:"wrap"}}>
            <span>📅 {a.date}</span>{a.presenter&&<span>👤 {a.presenter}</span>}{a.venue&&<span>📍 {a.venue}</span>}{a.college&&<span>🏛️ {a.college}</span>}{a.department&&<span>🏢 {a.department}</span>}<span>👥 {a.attendees||0}</span>
            {a.mediaStatus&&a.mediaStatus!=="Not contacted"&&<span style={{padding:"1px 6px",borderRadius:8,background:a.mediaStatus==="Published"?"#d5f5e3":a.mediaStatus==="Contacted"?"#fef9e7":"#eaf2f8",color:a.mediaStatus==="Published"?"#196f3d":a.mediaStatus==="Contacted"?"#7d6608":"#1a5276",fontSize:9,fontWeight:600}}>📺 {a.mediaStatus}{a.mediaChannel?` · ${a.mediaChannel}`:""}</span>}
          </div>
          {a.objectives&&<div style={{fontSize:11,color:"var(--color-text-tertiary)",marginTop:3}}>{a.objectives}</div>}
          {a.notes&&<div style={{fontSize:11,color:"#e67e22",marginTop:2,fontStyle:"italic"}}>📌 {a.notes}</div>}
          {a.mediaCoverageNotes&&<div style={{fontSize:11,color:"#8e44ad",marginTop:2}}>📺 {a.mediaCoverageNotes}</div>}
          {a.customFields&&a.customFields.length>0&&<div style={{marginTop:4,display:"flex",gap:6,flexWrap:"wrap"}}>{a.customFields.filter(f=>f.name).map((f,i)=><span key={i} style={{fontSize:10,padding:"2px 8px",borderRadius:8,background:"#f0e6f6",color:"#6c3483",border:"0.5px solid #d7bde2"}}>{f.name}: {f.value}</span>)}</div>}
        </div>
        <div style={{display:"flex",gap:3,flexShrink:0,alignItems:"flex-start"}}>
          <button onClick={()=>setShowExport(true)} title={t.exportReport} style={btnI}><Ic n="download" s={14} c="#2980b9"/></button>
          {canEdit&&<>
            {a.status!=="Completed"&&<button onClick={()=>onUpdate({status:"Completed"})} title={t.complete} style={btnI}><Ic n="check" s={14} c="#27ae60"/></button>}
            <button onClick={()=>sEd(true)} style={btnI}><Ic n="edit" s={14}/></button>
            <button onClick={onDelete} style={btnI}><Ic n="trash" s={14} c="#e74c3c"/></button>
          </>}
        </div>
      </div>
      {showExport&&<div onClick={()=>setShowExport(false)} style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.45)",display:"flex",alignItems:"center",justifyContent:"center",zIndex:999}}><div onClick={e=>e.stopPropagation()} style={{background:"#ffffff",borderRadius:14,padding:24,maxWidth:550,width:"92%",maxHeight:"80vh",overflow:"auto"}}><ExportModal t={t} activity={a} logoSrc={logoSrc||LOGO} pName={pName||""} pSub={pSub||""} onClose={()=>setShowExport(false)}/><button onClick={()=>setShowExport(false)} style={{...btnC,marginTop:14}}>{t.cancel}</button></div></div>}
    </div>
  );
}

// ─── ALL ACTIVITIES ───────────────────────────────────────────────────
function ActivitiesPage({t,allActs,logoSrc,pName,pSub,COLLEGES}){
  const [fS,sFS]=useState("All");const [fSt,sFSt]=useState("All");const [fC,sFC]=useState("All");const [fSearch,setFSearch]=useState("");
  const [showBulkExport,setShowBulkExport]=useState(false);
  const filtered=allActs.filter(a=>{
    if(fSearch){const q=fSearch.toLowerCase();if(!a.title.toLowerCase().includes(q)&&!(a.presenter||"").toLowerCase().includes(q)&&!(a.college||"").toLowerCase().includes(q)&&!(a.type||"").toLowerCase().includes(q))return false;}
    if(fS!=="All"&&a.source!==fS)return false;if(fSt!=="All"&&a.status!==fSt)return false;if(fC!=="All"&&a.college!==fC)return false;return true;
  });
  const exportCSV=()=>{
    const hdr="Title,Type,Date,Status,Source,Presenter,College,Attendees\n";
    const rows=filtered.map(a=>`"${a.title}","${a.type}","${a.date}","${a.status}","${a.source}","${a.presenter||""}","${a.college||""}",${a.attendees||0}`).join("\n");
    const blob=new Blob([hdr+rows],{type:"text/csv"});downloadFile(blob,"activities.csv");
  };
  return(
    <div>
      <div style={{display:"flex",justifyContent:"space-between",marginBottom:14}}>
        <div><h2 style={{fontSize:18,fontWeight:700,margin:0}}>{t.allActivities}</h2></div>
        <div style={{display:"flex",gap:8}}>
          <button onClick={()=>setShowBulkExport(true)} style={{...btnP,display:"flex",alignItems:"center",gap:4,fontSize:11,background:"#217346"}}><span>📊</span>{t.exportBulkExcel}</button>
          <button onClick={exportCSV} style={{...btnC,display:"flex",alignItems:"center",gap:4,fontSize:11}}><Ic n="download" s={14}/>{t.exportCSV}</button>
        </div>
      </div>
      <div style={{display:"flex",gap:8,marginBottom:12,flexWrap:"wrap",alignItems:"center"}}>
        <Flt l={t.source} v={fS} set={sFS} opts={[t.all,...SOURCES]}/>
        <Flt l={t.status} v={fSt} set={sFSt} opts={[t.all,...STATUSES]}/>
        <Flt l={t.college} v={fC} set={sFC} opts={[t.all,...COLLEGES,"All Colleges"]}/>
        <input value={fSearch} onChange={e=>setFSearch(e.target.value)} placeholder={"🔍 "+t.search+"..."} style={{...inp,width:"auto",minWidth:180,padding:"5px 10px",fontSize:11}}/>
        <span style={{fontSize:11,color:"var(--color-text-tertiary)"}}>{filtered.length} {t.results}</span>
      </div>
      {filtered.map(a=><ActCard key={a.id} t={t} a={a} canEdit={false} onUpdate={()=>{}} onDelete={()=>{}} logoSrc={logoSrc} pName={pName} pSub={pSub}/>)}
      {filtered.length===0&&<div style={{textAlign:"center",padding:50,color:"var(--color-text-tertiary)"}}>{t.noActivities}</div>}
      {showBulkExport&&<div onClick={()=>setShowBulkExport(false)} style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.45)",display:"flex",alignItems:"center",justifyContent:"center",zIndex:999}}><div onClick={e=>e.stopPropagation()} style={{background:"#ffffff",borderRadius:14,padding:24,maxWidth:550,width:"92%"}}><ExportModal t={t} activities={filtered} isBulk={true} logoSrc={logoSrc} pName={pName} pSub={pSub} onClose={()=>setShowBulkExport(false)}/><button onClick={()=>setShowBulkExport(false)} style={{...btnC,marginTop:14}}>{t.cancel}</button></div></div>}
    </div>
  );
}

// ─── INITIATIVES ──────────────────────────────────────────────────────
function InitiativesPage({t,inits,saveInits,isAdmin,trainers,logoSrc,pName,pSub,COLLEGES}){
  const [sf,sSf]=useState(false);
  const [fSource,setFSource]=useState("All");const [fStatus,setFStatus]=useState("All");const [fCollege,setFCollege]=useState("All");const [fSearch,setFSearch]=useState("");
  const add=d=>{saveInits([...inits,{id:gid(),...d,attendees:0,mediaNotified:false,mediaLinks:"",participants:"",college:d.college||"All Colleges"}]);sSf(false);};
  const upd=(id,u)=>saveInits(inits.map(i=>i.id===id?{...i,...u}:i));
  const del=id=>saveInits(inits.filter(i=>i.id!==id));
  const filtered=inits.filter(i=>{
    if(fSource!=="All"&&i.source!==fSource)return false;
    if(fStatus!=="All"&&i.status!==fStatus)return false;
    if(fCollege!=="All"&&i.college!==fCollege)return false;
    if(fSearch){const q=fSearch.toLowerCase();const s=[i.title,i.presenter,i.college,i.type,i.source,i.venue,i.objectives,i.notes].filter(Boolean).join(" ").toLowerCase();if(!s.includes(q))return false;}
    return true;
  });
  return(
    <div>
      <div style={{display:"flex",justifyContent:"space-between",marginBottom:12,flexWrap:"wrap",gap:8}}>
        <div><h2 style={{fontSize:18,fontWeight:700,margin:0}}>{t.initiatives}</h2><p style={{fontSize:12,color:"var(--color-text-secondary)",margin:"3px 0 0"}}>{t.initSubtitle}</p></div>
        {isAdmin&&<button onClick={()=>sSf(true)} style={{...btnP,display:"flex",alignItems:"center",gap:5}}><Ic n="plus" s={15} c="#fff"/>{t.add}</button>}
      </div>
      <div style={{display:"flex",gap:8,marginBottom:14,flexWrap:"wrap",alignItems:"center"}}>
        <select value={fSource} onChange={e=>setFSource(e.target.value)} style={{...inp,width:"auto",minWidth:140,padding:"5px 10px",fontSize:11}}>
          <option value="All">{t.all} {t.source}</option>
          {SOURCES.map(s=><option key={s}>{s}</option>)}
        </select>
        <select value={fStatus} onChange={e=>setFStatus(e.target.value)} style={{...inp,width:"auto",minWidth:120,padding:"5px 10px",fontSize:11}}>
          <option value="All">{t.all} {t.status}</option>
          {STATUSES.map(s=><option key={s}>{s}</option>)}
        </select>
        <select value={fCollege} onChange={e=>setFCollege(e.target.value)} style={{...inp,width:"auto",minWidth:160,padding:"5px 10px",fontSize:11}}>
          <option value="All">🏛️ {t.all} {t.college}</option>
          <option value="All Colleges">{t.all} {t.college} (university-wide)</option>
          {(COLLEGES||[]).map(c=><option key={c}>{c}</option>)}
        </select>
        <input value={fSearch} onChange={e=>setFSearch(e.target.value)} placeholder={"🔍 "+t.search+"..."} style={{...inp,width:"auto",minWidth:180,padding:"5px 10px",fontSize:11}}/>
        <span style={{fontSize:11,color:"var(--color-text-tertiary)"}}>{filtered.length}/{inits.length}</span>
      </div>
      {sf&&<ActForm t={t} onSave={add} onCancel={()=>sSf(false)} trainers={trainers} COLLEGES={COLLEGES}/>}
      {filtered.map(i=><ActCard key={i.id} t={t} a={i} canEdit={isAdmin} onUpdate={u=>upd(i.id,u)} onDelete={()=>del(i.id)} trainers={trainers} logoSrc={logoSrc} pName={pName} pSub={pSub} COLLEGES={COLLEGES}/>)}
      {filtered.length===0&&<div style={{textAlign:"center",padding:50,color:"var(--color-text-tertiary)"}}>{t.noActivities}</div>}
    </div>
  );
}

// ─── TRAINERS ─────────────────────────────────────────────────────────
function TrainersPage({t,trainers,saveTrainers,isAdmin,canEdit,COLLEGES,allActs,logoSrc,pName}){
  const [sf,sSf]=useState(false);const [eId,sEId]=useState(null);
  const [fCollege,setFCollege]=useState("All");const [fSearch,setFSearch]=useState("");
  const [bulkExtracting,setBE]=useState(false);const [bulkResults,setBR]=useState(null);const [bulkMsg,setBM]=useState("");
  const add=d=>{saveTrainers([...trainers,{id:gid(),...d}]);sSf(false);};
  const upd=d=>{saveTrainers(trainers.map(x=>x.id===eId?{...x,...d}:x));sEId(null);};
  const del=id=>saveTrainers(trainers.filter(x=>x.id!==id));
  const filteredTrainers=trainers.filter(x=>{
    if(fCollege!=="All"&&x.college!==fCollege)return false;
    if(fSearch){const q=fSearch.toLowerCase();const s=[x.name,x.email,x.phone,x.workplace,x.degree,x.generalSpec,x.specificSpec,x.college,...(x.linkedActivities||[])].filter(Boolean).join(" ").toLowerCase();if(!s.includes(q))return false;}
    return true;
  });
  const addBulk=items=>{const nw=items.map(i=>({id:gid(),name:i.name||"",dob:i.dob||"",gender:i.gender||"Male",phone:i.phone||"",email:i.email||"",workplace:i.workplace||"",degree:i.degree||"PhD",generalSpec:i.generalSpec||"",specificSpec:i.specificSpec||"",internal:i.internal??true,notes:i.notes||"",customFields:i.customFields||[],fileData:"",fileName:""}));saveTrainers([...trainers,...nw]);setBR(null);setBM("");};
  const handleBulkFile=async file=>{
    if(!file)return;setBE(true);setBM("");setBR(null);
    if(file.size>4*1024*1024){setBM("⚠️ Max 4MB");setBE(false);return;}
    const mtype=file.type||"";if(!mtype.startsWith("image/")&&mtype!=="application/pdf"&&!mtype.startsWith("text/")){setBM("⚠️ Use PDF, image, or camera");setBE(false);return;}
    try{const b64=await new Promise((r,j)=>{const rd=new FileReader();rd.onload=()=>r(rd.result.split(",")[1]);rd.onerror=()=>j();rd.readAsDataURL(file);});
      const items=await extractBulkTrainers(b64,mtype,t.dir==="rtl"?"ar":"en");
      if(items?.length>0){setBR(items);setBM(`✅ ${items.length} ${t.trainersFound}`);}else{setBM("❌ "+t.extractionError);}
    }catch{setBM("❌ "+t.extractionError);}setBE(false);
  };
  const dlTrainerFile=x=>{if(!x.fileData||!x.fileName)return;downloadFile(x.fileData,x.fileName);};

  return(
    <div>
      <div style={{display:"flex",justifyContent:"space-between",marginBottom:16}}>
        <div><h2 style={{fontSize:18,fontWeight:700,margin:0}}>{t.trainers}</h2><p style={{fontSize:12,color:"var(--color-text-secondary)",margin:"3px 0 0"}}>{t.trainerDB}</p></div>
        {(isAdmin||canEdit)&&<button onClick={()=>{sSf(true);sEId(null);}} style={{...btnP,display:"flex",alignItems:"center",gap:5}}><Ic n="plus" s={15} c="#fff"/>{t.newTrainer}</button>}
      </div>
      {(sf||eId)&&<TrForm t={t} initial={eId?trainers.find(x=>x.id===eId):null} onSave={eId?upd:add} onCancel={()=>{sSf(false);sEId(null);}} COLLEGES={COLLEGES} allActs={allActs}/>}

      {/* FILTERS */}
      <div style={{display:"flex",gap:8,marginBottom:14,flexWrap:"wrap",alignItems:"center"}}>
        <select value={fCollege} onChange={e=>setFCollege(e.target.value)} style={{...inp,width:"auto",minWidth:160,padding:"5px 10px",fontSize:11}}>
          <option value="All">🏛️ {t.all} {t.college}</option>
          {(COLLEGES||[]).map(c=><option key={c}>{c}</option>)}
        </select>
        <input value={fSearch} onChange={e=>setFSearch(e.target.value)} placeholder={"🔍 "+t.search+"..."} style={{...inp,width:"auto",minWidth:180,padding:"5px 10px",fontSize:11}}/>
        <span style={{fontSize:11,color:"var(--color-text-tertiary)"}}>{filteredTrainers.length}/{trainers.length}</span>
        <div style={{marginLeft:"auto",display:"flex",gap:6}}>
          <button onClick={()=>exportTrainersExcel(filteredTrainers,t)} style={{...btnP,padding:"5px 12px",fontSize:11,background:"#217346",display:"flex",alignItems:"center",gap:4}}>📊 {t.exportAsExcel}</button>
          <button onClick={()=>printTrainers(filteredTrainers,t,logoSrc,pName)} style={{...btnP,padding:"5px 12px",fontSize:11,background:"#c0392b",display:"flex",alignItems:"center",gap:4}}>🖨️ {t.exportAsPDF}</button>
        </div>
      </div>

      {/* BULK IMPORT */}
      {(isAdmin||canEdit)&&<div style={{background:"var(--color-background-secondary)",borderRadius:10,padding:14,marginBottom:14,border:"1px dashed var(--color-border-secondary)"}}>
        <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:6}}><span style={{fontSize:16}}>📋</span><div><div style={{fontSize:12,fontWeight:600}}>{t.bulkTrainerImport}</div><div style={{fontSize:10,color:"var(--color-text-tertiary)"}}>{t.bulkTrainerHint}</div></div></div>
        <div style={{display:"flex",gap:8,alignItems:"center",flexWrap:"wrap"}}>
          <label style={{...btnP,padding:"5px 12px",fontSize:11,cursor:"pointer",opacity:bulkExtracting?0.5:1}}><Ic n="plan" s={13} c="#fff"/>&nbsp;{t.fileUpload}<input type="file" accept="image/*,.pdf,.txt" onChange={e=>handleBulkFile(e.target.files[0])} style={{display:"none"}} disabled={bulkExtracting}/></label>
          <label style={{...btnP,padding:"5px 12px",fontSize:11,cursor:"pointer",background:"#27ae60",opacity:bulkExtracting?0.5:1}}>📷 {t.camera}<input type="file" accept="image/*" capture="environment" onChange={e=>handleBulkFile(e.target.files[0])} style={{display:"none"}} disabled={bulkExtracting}/></label>
          {bulkExtracting&&<span style={{fontSize:11,color:"#e67e22"}}>⏳ {t.extracting}</span>}
        </div>
        {bulkMsg&&<div style={{marginTop:6,fontSize:11,color:bulkResults?"#27ae60":"#e74c3c"}}>{bulkMsg}</div>}
        {bulkResults?.length>0&&<div style={{marginTop:10}}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:6}}><span style={{fontSize:11,fontWeight:600}}>{bulkResults.length} {t.trainersFound}:</span><button onClick={()=>addBulk(bulkResults)} style={{...btnP,padding:"4px 12px",fontSize:11}}>{t.importAll} ({bulkResults.length})</button></div>
          {bulkResults.map((tr,i)=><div key={i} style={{background:"#ffffff",borderRadius:6,padding:"6px 10px",marginBottom:4,border:"0.5px solid var(--color-border-tertiary)",fontSize:11,display:"flex",justifyContent:"space-between"}}><div><b>{tr.name||`Trainer ${i+1}`}</b> · {tr.degree} · {tr.specificSpec}</div><button onClick={()=>setBR(bulkResults.filter((_,idx)=>idx!==i))} style={{background:"none",border:"none",cursor:"pointer",color:"#e74c3c",fontSize:11}}>✕</button></div>)}
        </div>}
      </div>}

      <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(280px,1fr))",gap:12}}>
        {filteredTrainers.map(x=>(
          <div key={x.id} style={{...crd,padding:"14px 18px"}}>
            <div style={{display:"flex",justifyContent:"space-between",marginBottom:8}}>
              <div style={{display:"flex",gap:8,alignItems:"center"}}><div style={{width:36,height:36,borderRadius:"50%",background:"#0d324915",display:"flex",alignItems:"center",justifyContent:"center",fontSize:13,fontWeight:600,color:"#1a5276"}}>{(x.name||"?")[0]}</div><div><div style={{fontSize:13,fontWeight:600}}>{x.name}</div><div style={{fontSize:11,color:"var(--color-text-secondary)"}}>{x.degree} · {x.specificSpec}</div></div></div>
              {(isAdmin||canEdit)&&<div style={{display:"flex",gap:3}}><button onClick={()=>sEId(x.id)} style={btnI}><Ic n="edit" s={13}/></button><button onClick={()=>del(x.id)} style={btnI}><Ic n="trash" s={13} c="#e74c3c"/></button></div>}
            </div>
            <div style={{fontSize:11,color:"var(--color-text-secondary)",lineHeight:1.8}}>
              {x.email&&<div>📧 {x.email}</div>}{x.phone&&<div>📞 {x.phone}</div>}
              <div>🏢 {x.workplace||"-"} · {x.internal?"🟢 "+t.internal:"🔵 "+t.external}</div>
            </div>
            {x.college&&<div style={{marginTop:4}}><span style={{fontSize:9,padding:"2px 7px",borderRadius:8,background:"#e8d5f5",color:"#6c3483"}}>{x.college.replace("College of ","").replace("Department of ","")}</span></div>}
            {x.linkedActivities&&x.linkedActivities.length>0&&<div style={{marginTop:4,display:"flex",gap:4,flexWrap:"wrap"}}>{x.linkedActivities.map((la,i)=><span key={i} style={{fontSize:9,padding:"2px 6px",borderRadius:6,background:"#eaf2f8",color:"#1a5276"}}>{la}</span>)}</div>}
            {x.customFields?.length>0&&<div style={{marginTop:4,display:"flex",gap:4,flexWrap:"wrap"}}>{x.customFields.filter(f=>f.name).map((f,i)=><span key={i} style={{fontSize:9,padding:"2px 6px",borderRadius:6,background:"#f0e6f6",color:"#6c3483"}}>{f.name}: {f.value}</span>)}</div>}
            {x.fileName&&<div style={{marginTop:6}}><button onClick={()=>dlTrainerFile(x)} style={{display:"inline-flex",alignItems:"center",gap:4,padding:"3px 8px",borderRadius:5,border:"1px solid var(--color-border-tertiary)",background:"var(--color-background-secondary)",cursor:"pointer",fontSize:10,color:"#2980b9"}}><Ic n="download" s={12} c="#2980b9"/>{x.fileName}</button></div>}
          </div>
        ))}
      </div>
      {filteredTrainers.length===0&&!bulkResults&&<div style={{textAlign:"center",padding:50,color:"var(--color-text-tertiary)"}}>{t.noTrainers}</div>}
    </div>
  );
}

function TrForm({t,initial,onSave,onCancel,COLLEGES,allActs}){
  const [d,sD]=useState({name:initial?.name||"",dob:initial?.dob||"",gender:initial?.gender||"Male",phone:initial?.phone||"",email:initial?.email||"",workplace:initial?.workplace||"",degree:initial?.degree||"PhD",generalSpec:initial?.generalSpec||"",specificSpec:initial?.specificSpec||"",internal:initial?.internal??true,notes:initial?.notes||"",college:initial?.college||""});
  const [linkedActivities,setLA]=useState(initial?.linkedActivities||[]);
  const [customFields,setCF]=useState(initial?.customFields||[]);
  const [fileData,setFileData]=useState(initial?.fileData||"");
  const [fileName,setFileName]=useState(initial?.fileName||"");
  const [extracting,setExtracting]=useState(false);const [extractMsg,setExtractMsg]=useState("");
  const s=(k,v)=>sD({...d,[k]:v});
  const addCF=()=>setCF([...customFields,{id:gid(),name:"",value:""}]);
  const updCF=(id,f,v)=>setCF(customFields.map(x=>x.id===id?{...x,[f]:v}:x));
  const delCF=id=>setCF(customFields.filter(x=>x.id!==id));

  const handleFile=async file=>{
    if(!file)return;
    const {dataUrl:compressed, originalDataUrl}=await processFileForStorage(file);
    setFileData(compressed);setFileName(file.name);
    if(file.size>4*1024*1024){setExtractMsg("✅ File attached (compressed).");return;}
    const mtype=file.type||"";if(!mtype.startsWith("image/")&&mtype!=="application/pdf"&&!mtype.startsWith("text/")){setExtractMsg("✅ File attached. Use PDF/image for extraction.");return;}
    setExtracting(true);setExtractMsg("");
    try{
      const b64=originalDataUrl.split(",")[1];const result=await extractTrainerFromFile(b64,mtype,t.dir==="rtl"?"ar":"en");
      if(result){const newD={...d};["name","dob","gender","phone","email","workplace","degree","generalSpec","specificSpec","notes"].forEach(k=>{if(result[k]&&typeof result[k]==="string"&&result[k].trim())newD[k]=result[k].trim();});
        if(typeof result.internal==="boolean")newD.internal=result.internal;sD(newD);
        if(result.customFields?.length>0)setCF(prev=>[...prev,...result.customFields.map(cf=>({id:gid(),name:cf.name||"",value:cf.value||""}))]);
        setExtractMsg(t.extractionDone);
      }else{setExtractMsg("✅ File attached. Fill manually.");}
    }catch{setExtractMsg("✅ File attached. Extraction error.");}setExtracting(false);
  };

  const doSave=()=>{if(!d.name)return;onSave({...d,linkedActivities,customFields:customFields.filter(f=>f.name.trim()),fileData,fileName});};

  return(
    <div style={{...crd,border:"1.5px solid #c9a84c",marginBottom:16}}>
      <h3 style={{fontSize:14,fontWeight:600,margin:"0 0 8px"}}>{initial?t.edit:t.newTrainer}</h3>
      {/* FILE UPLOAD */}
      {!initial&&<div style={{background:"var(--color-background-secondary)",borderRadius:10,padding:14,marginBottom:12,border:"1px dashed var(--color-border-secondary)"}}>
        <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:6}}><span style={{fontSize:16}}>📄</span><div><div style={{fontSize:12,fontWeight:600}}>{t.extractFromFile}</div><div style={{fontSize:10,color:"var(--color-text-tertiary)"}}>{t.trainerExtractHint}</div></div></div>
        <div style={{display:"flex",gap:8,alignItems:"center",flexWrap:"wrap"}}>
          <label style={{...btnP,padding:"5px 12px",fontSize:11,cursor:"pointer",opacity:extracting?0.5:1}}><Ic n="plan" s={13} c="#fff"/>&nbsp;{t.fileUpload}<input type="file" accept="image/*,.pdf,.txt" onChange={e=>handleFile(e.target.files[0])} style={{display:"none"}} disabled={extracting}/></label>
          <label style={{...btnP,padding:"5px 12px",fontSize:11,cursor:"pointer",background:"#27ae60",opacity:extracting?0.5:1}}>📷 {t.camera}<input type="file" accept="image/*" capture="environment" onChange={e=>handleFile(e.target.files[0])} style={{display:"none"}} disabled={extracting}/></label>
          {extracting&&<span style={{fontSize:11,color:"#e67e22"}}>⏳ {t.extracting}</span>}
          {fileName&&<span style={{fontSize:11,color:"#27ae60"}}>📎 {fileName}</span>}
        </div>
        {extractMsg&&<div style={{marginTop:6,fontSize:11,color:extractMsg.includes("✅")||extractMsg===t.extractionDone?"#27ae60":"#e74c3c"}}>{extractMsg}</div>}
      </div>}
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:10}}>
        <div style={{gridColumn:"1/-1"}}><label style={lbl}>{t.name}</label><input value={d.name} onChange={e=>s("name",e.target.value)} style={inp}/></div>
        <div><label style={lbl}>{t.dob}</label><input type="date" value={d.dob} onChange={e=>s("dob",e.target.value)} style={inp}/></div>
        <div><label style={lbl}>{t.gender}</label><select value={d.gender} onChange={e=>s("gender",e.target.value)} style={inp}><option>{t.male}</option><option>{t.female}</option></select></div>
        <div><label style={lbl}>{t.phone}</label><input value={d.phone} onChange={e=>s("phone",e.target.value)} style={inp}/></div>
        <div><label style={lbl}>{t.email}</label><input value={d.email} onChange={e=>s("email",e.target.value)} style={inp}/></div>
        <div><label style={lbl}>{t.workplace}</label><input value={d.workplace} onChange={e=>s("workplace",e.target.value)} style={inp}/></div>
        <div><label style={lbl}>{t.degree}</label><select value={d.degree} onChange={e=>s("degree",e.target.value)} style={inp}><option>PhD</option><option>MSc</option><option>BSc</option><option>Board Certified</option></select></div>
        <div><label style={lbl}>{t.generalSpec}</label><input value={d.generalSpec} onChange={e=>s("generalSpec",e.target.value)} style={inp}/></div>
        <div><label style={lbl}>{t.specificSpec}</label><input value={d.specificSpec} onChange={e=>s("specificSpec",e.target.value)} style={inp}/></div>
        <div><label style={lbl}>{t.internal}/{t.external}</label><select value={d.internal?"true":"false"} onChange={e=>s("internal",e.target.value==="true")} style={inp}><option value="true">{t.internal}</option><option value="false">{t.external}</option></select></div>
        <div><label style={lbl}>{t.college}</label><select value={d.college} onChange={e=>s("college",e.target.value)} style={inp}><option value="">— {t.all} —</option>{(COLLEGES||[]).map(c=><option key={c}>{c}</option>)}</select></div>
        <div style={{gridColumn:"span 2"}}><label style={lbl}>{t.notes}</label><input value={d.notes} onChange={e=>s("notes",e.target.value)} style={inp}/></div>
        {/* LINKED ACTIVITIES */}
        <div style={{gridColumn:"1/-1",borderTop:"1px dashed var(--color-border-tertiary)",paddingTop:8,marginTop:4}}>
          <label style={lbl}>🔗 Linked activities / workshops</label>
          <select onChange={e=>{if(e.target.value&&!linkedActivities.includes(e.target.value)){setLA([...linkedActivities,e.target.value]);}e.target.value="";}} style={{...inp,marginBottom:6}}>
            <option value="">— Select activity to link —</option>
            {(allActs||[]).filter(a=>!linkedActivities.includes(a.title)&&(!d.college||!a.college||a.college===d.college)).map(a=><option key={a.id} value={a.title}>{a.title} ({a.date||""}){a.college?` · ${a.college.replace("College of ","").replace("Department of ","")}`:""}</option>)}
          </select>
          {linkedActivities.length>0&&<div style={{display:"flex",gap:4,flexWrap:"wrap"}}>{linkedActivities.map((la,i)=>(
            <span key={i} style={{fontSize:10,padding:"3px 8px",borderRadius:6,background:"#eaf2f8",color:"#1a5276",display:"inline-flex",alignItems:"center",gap:4}}>{la}<button onClick={()=>setLA(linkedActivities.filter((_,idx)=>idx!==i))} style={{background:"none",border:"none",cursor:"pointer",color:"#e74c3c",fontSize:10,padding:0}}>✕</button></span>
          ))}</div>}
        </div>
        {/* CUSTOM FIELDS */}
        {customFields.length>0&&<div style={{gridColumn:"1/-1",borderTop:"1px dashed var(--color-border-tertiary)",paddingTop:8}}><span style={{fontSize:12,fontWeight:600,color:"#8e44ad"}}>📋 {t.extractedFields}</span></div>}
        {customFields.map(cf=>(
          <div key={cf.id} style={{gridColumn:"1/-1",display:"flex",gap:8,alignItems:"flex-start"}}>
            <div style={{flex:1}}><label style={lbl}>{t.fieldName}</label><input value={cf.name} onChange={e=>updCF(cf.id,"name",e.target.value)} style={inp}/></div>
            <div style={{flex:2}}><label style={lbl}>{t.fieldValue}</label><input value={cf.value} onChange={e=>updCF(cf.id,"value",e.target.value)} style={inp}/></div>
            <button onClick={()=>delCF(cf.id)} style={{...btnI,marginTop:18}}><Ic n="trash" s={14} c="#e74c3c"/></button>
          </div>
        ))}
        <div style={{gridColumn:"1/-1"}}><button onClick={addCF} style={{...btnC,fontSize:11,padding:"5px 12px",display:"inline-flex",alignItems:"center",gap:4}}><Ic n="plus" s={13}/>{t.addCustomField}</button></div>
      </div>
      <div style={{display:"flex",gap:8,marginTop:12,justifyContent:"flex-end"}}><button onClick={onCancel} style={btnC}>{t.cancel}</button><button onClick={doSave} style={{...btnP,opacity:d.name?1:0.5}}>{t.save}</button></div>
    </div>
  );
}

// ─── DIRECTIVES PAGE ─────────────────────────────────────────────────
function DirectivesPage({t,directives,saveDirectives,allActs,inits,saveInits,plans,savePlans,COLLEGES,isAdmin,logoSrc,pName}){
  const [sf,sSf]=useState(false);const [editId,setEditId]=useState(null);const [fStatus,setFStatus]=useState("All");const [fSearch,setFSearch]=useState("");
  const [expandedId,setExpandedId]=useState(null);const [linkingCol,setLinkingCol]=useState(null);
  const [showExport,setShowExport]=useState(false);const [expSec,setExpSec]=useState({summary:true,details:true,compliance:true,overdue:true});
  const now=new Date();

  const add=d=>{saveDirectives([...directives,{id:gid(),createdDate:new Date().toISOString().split("T")[0],...d,compliance:{}}]);sSf(false);};
  const upd=d=>{saveDirectives(directives.map(x=>x.id===editId?{...x,...d}:x));setEditId(null);};
  const del=id=>saveDirectives(directives.filter(x=>x.id!==id));

  // Auto-match activities to directives
  const getCompliance=(dir)=>{
    const targets=dir.targetColleges||[];
    const isAll=targets.includes("All")||targets.length===0;
    const colleges=isAll?COLLEGES:targets;
    const manual=dir.compliance||{};

    return colleges.map(col=>{
      // Check manual compliance first
      if(manual[col]?.status==="Compliant")return{college:col,...manual[col]};
      // Auto-detect: look for matching activities in this college
      const colActs=allActs.filter(a=>a.college===col);
      // Match by: directive title words in activity title, or same type, or linked activityId
      const linked=manual[col]?.activityId?colActs.find(a=>a.id===manual[col].activityId):null;
      if(linked)return{college:col,status:linked.status==="Completed"?"Compliant":linked.status==="Cancelled"?"Not started":"In progress",activityId:linked.id,activityTitle:linked.title,autoMatched:false};
      // Fuzzy match by title keywords
      const dirWords=dir.title.toLowerCase().split(/\s+/).filter(w=>w.length>3);
      const matched=colActs.find(a=>{
        const aTitle=a.title.toLowerCase();
        return dirWords.filter(w=>aTitle.includes(w)).length>=Math.max(1,Math.floor(dirWords.length/2));
      });
      if(matched)return{college:col,status:matched.status==="Completed"?"Compliant":matched.status==="Cancelled"?"Not started":"In progress",activityId:matched.id,activityTitle:matched.title,autoMatched:true};
      return{college:col,status:"Not started",activityId:null,activityTitle:null};
    });
  };

  const setColCompliance=(dirId,college,data)=>{
    saveDirectives(directives.map(d=>{
      if(d.id!==dirId)return d;
      return{...d,compliance:{...d.compliance,[college]:{...d.compliance?.[college],...data}}};
    }));
  };

  const linkActivityToCollege=(dirId,college,actId)=>{
    const act=allActs.find(a=>a.id===actId);
    if(!act)return;
    setColCompliance(dirId,college,{activityId:actId,activityTitle:act.title,status:act.status==="Completed"?"Compliant":"In progress"});
    setLinkingCol(null);
  };

  // Create initiative for each target college
  const pushToColleges=(dir)=>{
    const targets=dir.targetColleges||[];
    const isAll=targets.includes("All")||targets.length===0;
    const colleges=isAll?COLLEGES:targets;
    const comp=getCompliance(dir);
    const needsInit=comp.filter(c=>c.status==="Not started").map(c=>c.college);
    if(needsInit.length===0)return;
    const newInits=needsInit.map(col=>({
      id:gid(),title:dir.title,type:dir.activityType||"Workshop",date:dir.deadline||"",presenter:"",delivery:"In-Person",target:"Faculty",venue:"",duration:"",objectives:dir.description||"",notes:`${t.directiveSource}: ${dir.title} (${dir.directiveType||""})`,status:"Planned",source:dir.directiveType||"Department Initiative",department:"",college:col,
      attendees:0,mediaNotified:false,mediaLinks:"",participants:"",customFields:[]
    }));
    saveInits([...inits,...newInits]);
    // Link them in compliance
    const updDir={...dir,compliance:{...dir.compliance}};
    newInits.forEach(ni=>{updDir.compliance[ni.college]={activityId:ni.id,activityTitle:ni.title,status:"In progress"};});
    saveDirectives(directives.map(d=>d.id===dir.id?updDir:d));
  };

  const filtered=directives.filter(d=>{
    if(fStatus==="active"&&getCompliance(d).every(c=>c.status==="Compliant"))return false;
    if(fStatus==="completed"&&!getCompliance(d).every(c=>c.status==="Compliant"))return false;
    if(fSearch){const q=fSearch.toLowerCase();if(!d.title.toLowerCase().includes(q)&&!(d.description||"").toLowerCase().includes(q))return false;}
    return true;
  });

  const exportDirectivesExcel=()=>{
    const wb=XLSX.utils.book_new();
    const sumData=filtered.map(dir=>{const comp=getCompliance(dir);const c=comp.filter(x=>x.status==="Compliant").length;return{"Title":dir.title,"Type":dir.directiveType||"","Deadline":dir.deadline||"","Colleges":comp.length,"Compliant":c,"In progress":comp.filter(x=>x.status==="In progress").length,"Not started":comp.filter(x=>x.status==="Not started").length,"Rate":comp.length?Math.round(c/comp.length*100)+"%":"0%","Created":dir.createdDate||""};});
    if(expSec.summary)XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(sumData),"Summary");
    if(expSec.compliance){const rows=[];filtered.forEach(dir=>{getCompliance(dir).forEach(c=>{rows.push({"Directive":dir.title,"College":c.college,"Status":c.status,"Linked activity":c.activityTitle||"","Auto-matched":c.autoMatched?"Yes":"No"});});});XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(rows),"Compliance Detail");}
    if(expSec.overdue){const od=filtered.filter(d=>d.deadline&&new Date(d.deadline)<now).map(d=>{const comp=getCompliance(d);return{"Title":d.title,"Deadline":d.deadline,"Days overdue":Math.ceil((now-new Date(d.deadline))/86400000),"Compliance":Math.round(comp.filter(c=>c.status==="Compliant").length/comp.length*100)+"%","Not started colleges":comp.filter(c=>c.status==="Not started").map(c=>c.college.replace("College of ","")).join(", ")};});if(od.length>0)XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(od),"Overdue");}
    if(wb.SheetNames.length===0)XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet([{Note:"No sections selected"}]),"Info");
    XLSX.writeFile(wb,`Directives_${new Date().toISOString().split('T')[0]}.xlsx`);
  };

  const printDirectives=()=>{
    const isRtl=t.dir==="rtl";const al=isRtl?"right":"left";
    let html=`<html dir="${isRtl?'rtl':'ltr'}"><head><meta charset="utf-8"><style>@import url('https://fonts.googleapis.com/css2?family=Noto+Sans+Arabic:wght@400;700&display=swap');body{font-family:'Noto Sans Arabic','Noto Sans',sans-serif;padding:30px;font-size:12px;}h1{color:#1a5276;font-size:20px;}h2{color:#1a5276;font-size:15px;border-bottom:2px solid #1a5276;padding-bottom:4px;}table{width:100%;border-collapse:collapse;margin:8px 0 16px;}th,td{border:1px solid #ddd;padding:5px 8px;text-align:${al};font-size:11px;}th{background:#f4f6f7;font-weight:700;}.badge{padding:2px 8px;border-radius:8px;font-size:10px;font-weight:600;}</style></head><body>`;
    if(logoSrc)html+=`<div style="display:flex;align-items:center;gap:14px;margin-bottom:16px;"><img src="${logoSrc}" width="55" style="border-radius:50%;background:#0d3249;"/><div><h1 style="margin:0">${pName||""}</h1><p style="margin:2px 0;font-size:12px;color:#6c7a89;">${t.directivesPage} — ${new Date().toLocaleDateString()}</p></div></div>`;
    if(expSec.summary){html+=`<h2>${t.directivesPage} (${filtered.length})</h2><table><tr><th>#</th><th>Title</th><th>Type</th><th>Deadline</th><th>Colleges</th><th>Compliance</th></tr>`;filtered.forEach((d,i)=>{const comp=getCompliance(d);const pct=comp.length?Math.round(comp.filter(c=>c.status==="Compliant").length/comp.length*100):0;html+=`<tr><td>${i+1}</td><td>${d.title}</td><td>${d.directiveType||""}</td><td>${d.deadline||""}</td><td>${comp.length}</td><td><span class="badge" style="background:${pct>=80?"#d5f5e3":pct>=50?"#fef5e7":"#fdedec"};color:${pct>=80?"#196f3d":pct>=50?"#b7950b":"#c0392b"}">${pct}%</span></td></tr>`;});html+=`</table>`;}
    if(expSec.compliance){filtered.forEach(d=>{const comp=getCompliance(d);html+=`<h2>${d.title}</h2><table><tr><th>College</th><th>Status</th><th>Linked activity</th></tr>`;comp.forEach(c=>{html+=`<tr><td>${c.college.replace("College of ","").replace("Department of ","")}</td><td style="color:${c.status==="Compliant"?"#27ae60":c.status==="In progress"?"#b7950b":"#e74c3c"};font-weight:700">${c.status}</td><td>${c.activityTitle||"—"}</td></tr>`;});html+=`</table>`;});}
    html+=`</body></html>`;showPrintPreview(html);
  };

  return(
    <div>
      <div style={{display:"flex",justifyContent:"space-between",marginBottom:12,flexWrap:"wrap",gap:8}}>
        <div><h2 style={{fontSize:18,fontWeight:700,margin:0}}>{t.directivesPage}</h2><p style={{fontSize:12,color:"var(--color-text-secondary)",margin:"3px 0 0"}}>{t.directivesSubtitle}</p></div>
        {isAdmin&&<div style={{display:"flex",gap:6}}>
          <button onClick={()=>setShowExport(!showExport)} style={{...btnP,padding:"6px 14px",fontSize:11,background:showExport?"#7f8c8d":"#2980b9"}}>📤 Export</button>
          <button onClick={()=>sSf(true)} style={{...btnP,display:"flex",alignItems:"center",gap:5}}><Ic n="plus" s={15} c="#fff"/>{t.newDirective}</button>
        </div>}
      </div>
      {showExport&&<div style={{...crd,border:"2px solid #2980b9",marginBottom:14,padding:14}}>
        <div style={{display:"flex",gap:6,marginBottom:10,flexWrap:"wrap"}}>
          {[{k:"summary",l:"📊 Summary"},{k:"compliance",l:"🏛️ Compliance detail"},{k:"overdue",l:"⚠️ Overdue"}].map(s=>(
            <label key={s.k} style={{display:"flex",alignItems:"center",gap:4,fontSize:11,padding:"4px 10px",borderRadius:6,background:expSec[s.k]?"#eaf2f8":"#f5f5f5",border:`1px solid ${expSec[s.k]?"#2980b9":"#ddd"}`,cursor:"pointer"}}>
              <input type="checkbox" checked={expSec[s.k]} onChange={()=>setExpSec({...expSec,[s.k]:!expSec[s.k]})} style={{accentColor:"#1a5276"}}/>{s.l}
            </label>))}
        </div>
        <div style={{display:"flex",gap:8}}>
          <button onClick={()=>{exportDirectivesExcel();setShowExport(false);}} style={{...btnP,padding:"6px 14px",fontSize:11,background:"#217346"}}>📊 Excel</button>
          <button onClick={()=>{printDirectives();setShowExport(false);}} style={{...btnP,padding:"6px 14px",fontSize:11,background:"#c0392b"}}>🖨️ Print / PDF</button>
        </div>
      </div>}
      <div style={{display:"flex",gap:8,marginBottom:14,flexWrap:"wrap",alignItems:"center"}}>
        <select value={fStatus} onChange={e=>setFStatus(e.target.value)} style={{...inp,width:"auto",minWidth:130,padding:"5px 10px",fontSize:11}}>
          <option value="All">{t.all}</option>
          <option value="active">🔄 Active</option>
          <option value="completed">✅ Completed</option>
        </select>
        <input value={fSearch} onChange={e=>setFSearch(e.target.value)} placeholder={"🔍 "+t.search+"..."} style={{...inp,width:"auto",minWidth:180,padding:"5px 10px",fontSize:11}}/>
        <span style={{fontSize:11,color:"var(--color-text-tertiary)"}}>{filtered.length}/{directives.length}</span>
      </div>

      {sf&&<DirectiveForm t={t} onSave={add} onCancel={()=>sSf(false)} COLLEGES={COLLEGES}/>}
      {editId&&<DirectiveForm t={t} onSave={upd} onCancel={()=>setEditId(null)} COLLEGES={COLLEGES} initial={directives.find(d=>d.id===editId)}/>}

      {filtered.map(dir=>{
        const comp=getCompliance(dir);
        const compliant=comp.filter(c=>c.status==="Compliant").length;
        const inProg=comp.filter(c=>c.status==="In progress").length;
        const notStarted=comp.filter(c=>c.status==="Not started").length;
        const pct=comp.length?Math.round(compliant/comp.length*100):0;
        const isOverdue=dir.deadline&&new Date(dir.deadline)<now&&pct<100;
        const expanded=expandedId===dir.id;

        return(
          <div key={dir.id} style={{...crd,marginBottom:12,border:isOverdue?"1.5px solid #e74c3c":"0.5px solid var(--color-border-tertiary)"}}>
            {/* HEADER */}
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",gap:8,cursor:"pointer"}} onClick={()=>setExpandedId(expanded?null:dir.id)}>
              <div style={{flex:1,minWidth:0}}>
                <div style={{display:"flex",gap:6,alignItems:"center",flexWrap:"wrap",marginBottom:4}}>
                  <span style={{fontSize:15,fontWeight:700}}>{dir.title}</span>
                  {dir.directiveType&&<span style={{fontSize:9,padding:"2px 7px",borderRadius:8,background:"#fef5e7",color:"#b7950b"}}>{dir.directiveType}</span>}
                  {isOverdue&&<span style={{fontSize:9,padding:"2px 7px",borderRadius:8,background:"#fdedec",color:"#e74c3c",fontWeight:700}}>⚠️ {t.overdue}</span>}
                  {pct===100&&<span style={{fontSize:9,padding:"2px 7px",borderRadius:8,background:"#d5f5e3",color:"#196f3d",fontWeight:700}}>✅ 100%</span>}
                </div>
                <div style={{fontSize:11,color:"var(--color-text-secondary)"}}>{t.deadline}: {dir.deadline||"—"} · {t.targetColleges}: {(dir.targetColleges||[]).includes("All")?t.allColleges:comp.length} · Created: {dir.createdDate}</div>
                {/* Progress bar */}
                <div style={{marginTop:6,display:"flex",gap:8,alignItems:"center"}}>
                  <div style={{flex:1,height:8,background:"#f0f0f0",borderRadius:4,overflow:"hidden",maxWidth:300}}>
                    <div style={{height:"100%",width:`${pct}%`,background:pct===100?"#27ae60":pct>50?"#f39c12":"#e74c3c",borderRadius:4,transition:"width 0.3s"}}/>
                  </div>
                  <span style={{fontSize:11,fontWeight:600,color:pct===100?"#27ae60":"#1a5276"}}>{pct}%</span>
                  <span style={{fontSize:10,color:"var(--color-text-tertiary)"}}>✅{compliant} 🔄{inProg} ⬜{notStarted}</span>
                </div>
              </div>
              <div style={{display:"flex",gap:3,flexShrink:0}}>
                {isAdmin&&<button onClick={e=>{e.stopPropagation();setEditId(dir.id);}} style={btnI}><Ic n="edit" s={13}/></button>}
                {isAdmin&&<button onClick={e=>{e.stopPropagation();del(dir.id);}} style={btnI}><Ic n="trash" s={13} c="#e74c3c"/></button>}
                <span style={{fontSize:16,color:"var(--color-text-tertiary)",padding:"0 4px"}}>{expanded?"▲":"▼"}</span>
              </div>
            </div>

            {/* EXPANDED: College compliance */}
            {expanded&&<div style={{marginTop:12,borderTop:"1px solid var(--color-border-tertiary)",paddingTop:12}}>
              {dir.description&&<div style={{fontSize:12,color:"var(--color-text-secondary)",marginBottom:10,padding:"8px 12px",background:"var(--color-background-secondary)",borderRadius:6}}>{dir.description}</div>}
              {/* Push button */}
              {isAdmin&&comp.some(c=>c.status==="Not started")&&<button onClick={()=>pushToColleges(dir)} style={{...btnP,padding:"6px 14px",fontSize:11,marginBottom:10,background:"#e67e22",display:"flex",alignItems:"center",gap:4}}>⚡ {t.createForColleges} ({comp.filter(c=>c.status==="Not started").length})</button>}
              {/* College rows */}
              {comp.map((c,i)=>(
                <div key={i} style={{display:"flex",alignItems:"center",gap:8,padding:"8px 10px",borderBottom:i<comp.length-1?"0.5px solid var(--color-border-tertiary)":"none",fontSize:12}}>
                  <div style={{width:8,height:8,borderRadius:"50%",background:c.status==="Compliant"?"#27ae60":c.status==="In progress"?"#f39c12":"#e74c3c",flexShrink:0}}/>
                  <span style={{flex:1,fontWeight:500}}>{c.college.replace("College of ","").replace("Department of ","")}</span>
                  <span style={{fontSize:10,padding:"2px 8px",borderRadius:8,background:c.status==="Compliant"?"#d5f5e3":c.status==="In progress"?"#fef5e7":"#fdedec",color:c.status==="Compliant"?"#196f3d":c.status==="In progress"?"#b7950b":"#c0392b",fontWeight:600}}>{c.status==="Compliant"?t.compliant:c.status==="In progress"?t.inProgress:t.notCompliant}</span>
                  {c.activityTitle&&<span style={{fontSize:10,color:"#2980b9",maxWidth:200,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}} title={c.activityTitle}>📋 {c.activityTitle}{c.autoMatched?" (auto)":""}</span>}
                  {isAdmin&&c.status!=="Compliant"&&<button onClick={()=>setLinkingCol({directiveId:dir.id,college:c.college})} style={{...btnC,padding:"2px 8px",fontSize:10}}>{t.linkActivity}</button>}
                  {isAdmin&&c.status!=="Compliant"&&<button onClick={()=>setColCompliance(dir.id,c.college,{status:"Compliant",manualDate:new Date().toISOString().split("T")[0]})} style={{...btnC,padding:"2px 8px",fontSize:10,color:"#27ae60",borderColor:"#27ae60"}}>{t.markCompliant}</button>}
                </div>
              ))}
            </div>}
          </div>
        );
      })}
      {filtered.length===0&&<div style={{textAlign:"center",padding:50,color:"var(--color-text-tertiary)"}}>{directives.length===0?"No directives yet":"No matches"}</div>}

      {/* LINK ACTIVITY MODAL */}
      {linkingCol&&<div onClick={()=>setLinkingCol(null)} style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.45)",display:"flex",alignItems:"center",justifyContent:"center",zIndex:999}}>
        <div onClick={e=>e.stopPropagation()} style={{background:"#ffffff",borderRadius:14,padding:24,maxWidth:500,width:"92%",maxHeight:"80vh",overflow:"auto"}}>
          <h3 style={{fontSize:14,fontWeight:600,margin:"0 0 12px"}}>{t.linkActivity} — {linkingCol.college.replace("College of ","").replace("Department of ","")}</h3>
          <div style={{maxHeight:300,overflowY:"auto"}}>
            {allActs.filter(a=>a.college===linkingCol.college).map(a=>(
              <div key={a.id} onClick={()=>linkActivityToCollege(linkingCol.directiveId,linkingCol.college,a.id)} style={{display:"flex",alignItems:"center",gap:8,padding:"8px 10px",borderBottom:"0.5px solid var(--color-border-tertiary)",cursor:"pointer",fontSize:12,borderRadius:6,":hover":{background:"#eaf2f8"}}}>
                <div style={{width:6,height:6,borderRadius:"50%",background:sC(a.status),flexShrink:0}}/>
                <div style={{flex:1}}><div style={{fontWeight:500}}>{a.title}</div><div style={{fontSize:10,color:"var(--color-text-tertiary)"}}>{a.date} · {a.type} · {a.status}</div></div>
              </div>
            ))}
            {allActs.filter(a=>a.college===linkingCol.college).length===0&&<div style={{fontSize:12,color:"var(--color-text-tertiary)",padding:20,textAlign:"center"}}>No activities for this college</div>}
          </div>
          <button onClick={()=>setLinkingCol(null)} style={{...btnC,marginTop:12}}>{t.cancel}</button>
        </div>
      </div>}
    </div>
  );
}

function DirectiveForm({t,onSave,onCancel,COLLEGES,initial}){
  const [d,sD]=useState({title:initial?.title||"",description:initial?.description||"",directiveType:initial?.directiveType||"Department Initiative",activityType:initial?.activityType||"Workshop",deadline:initial?.deadline||"",targetColleges:initial?.targetColleges||["All"]});
  const [selAll,setSelAll]=useState((initial?.targetColleges||["All"]).includes("All"));
  const s=(k,v)=>sD({...d,[k]:v});
  const toggleCollege=(col)=>{
    if(selAll)return;
    const cur=d.targetColleges||[];
    sD({...d,targetColleges:cur.includes(col)?cur.filter(c=>c!==col):[...cur,col]});
  };
  const toggleAll=()=>{
    setSelAll(!selAll);
    sD({...d,targetColleges:!selAll?["All"]:[]});
  };

  return(
    <div style={{...crd,border:"1.5px solid #c9a84c",marginBottom:14}}>
      <h3 style={{fontSize:14,fontWeight:600,margin:"0 0 10px"}}>{initial?t.edit:t.newDirective}</h3>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:12}}>
        <div style={{gridColumn:"1/-1"}}><label style={lbl}>{t.directiveTitle}</label><input value={d.title} onChange={e=>s("title",e.target.value)} style={inp} placeholder="e.g. Cybersecurity awareness workshop for all colleges"/></div>
        <div><label style={lbl}>{t.directiveType}</label><select value={d.directiveType} onChange={e=>s("directiveType",e.target.value)} style={inp}><option>Department Initiative</option><option>University Initiative</option><option>Ministry Directive</option><option>College Request</option></select></div>
        <div><label style={lbl}>{t.type}</label><select value={d.activityType} onChange={e=>s("activityType",e.target.value)} style={inp}>{ACT_TYPES.map(at=><option key={at}>{at}</option>)}</select></div>
        <div><label style={lbl}>{t.deadline}</label><input type="date" value={d.deadline} onChange={e=>s("deadline",e.target.value)} style={inp}/></div>
        <div style={{gridColumn:"1/-1"}}><label style={lbl}>{t.directiveDesc}</label><textarea value={d.description} onChange={e=>s("description",e.target.value)} rows={2} style={{...inp,resize:"vertical"}}/></div>
      </div>
      {/* College selection */}
      <div style={{marginBottom:12}}>
        <label style={lbl}>{t.targetColleges}</label>
        <label style={{display:"flex",alignItems:"center",gap:6,marginBottom:8,fontSize:12,cursor:"pointer",fontWeight:600}}>
          <input type="checkbox" checked={selAll} onChange={toggleAll} style={{accentColor:"#1a5276"}}/>
          🏛️ {t.allColleges} ({(COLLEGES||[]).length})
        </label>
        {!selAll&&<div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(200px,1fr))",gap:4,maxHeight:150,overflowY:"auto",padding:"8px",border:"1px solid var(--color-border-tertiary)",borderRadius:8}}>
          {(COLLEGES||[]).map(c=>(
            <label key={c} style={{display:"flex",alignItems:"center",gap:6,fontSize:11,cursor:"pointer",padding:"3px 6px",borderRadius:4,background:d.targetColleges.includes(c)?"#eaf2f8":"transparent"}}>
              <input type="checkbox" checked={d.targetColleges.includes(c)} onChange={()=>toggleCollege(c)} style={{accentColor:"#1a5276"}}/>
              {c.replace("College of ","").replace("Department of ","")}
            </label>
          ))}
        </div>}
      </div>
      <div style={{display:"flex",gap:8,justifyContent:"flex-end"}}>
        <button onClick={onCancel} style={btnC}>{t.cancel}</button>
        <button onClick={()=>d.title.trim()&&onSave(d)} style={{...btnP,opacity:d.title.trim()?1:0.5}}>{t.save}</button>
      </div>
    </div>
  );
}

// ─── STAFF PAGE ──────────────────────────────────────────────────────
function StaffPage({t,allActs,corr,trainers,saveTrainers,COLLEGES,isAdmin,staffRoster,saveRoster,staffMerges,saveStaffMerges,logoSrc,pName,reports}){
  const [mergeFor,setMergeFor]=useState(null);
  // Build alias→canonical lookup from staffMerges
  const aliasMap=useMemo(()=>{
    const m={};
    (staffMerges||[]).forEach(g=>{(g.aliases||[]).forEach(a=>{m[(a||"").trim().toLowerCase()]=g.canonical;});});
    return m;
  },[staffMerges]);
  const resolveCanonical=(name)=>{
    if(!name)return name;
    return aliasMap[name.trim().toLowerCase()]||name;
  };
  const [fCollege,setFCollege]=useState("All");const [fSearch,setFSearch]=useState("");const [fStatus,setFStatus]=useState("All");
  const [viewAct,setViewAct]=useState(null);const [letterForm,setLetterForm]=useState(null);
  const [showAddForm,setShowAddForm]=useState(false);const [importMsg,setImportMsg]=useState("");
  const now=new Date();

  // Helper: normalize name for fuzzy matching
  const normName=(n)=>{
    if(!n)return"";
    return n.trim().replace(/^(د\.|أ\.د\.|أ\.م\.|م\.|م\.د\.|م\.م\.|Prof\.|Dr\.|Asst\.\s*Prof\.|Assoc\.\s*Prof\.|Mr\.|Mrs\.|Ms\.|Eng\.|Lect\.)\s*/gi,"").replace(/\s+/g," ").trim().toLowerCase();
  };
  const fuzzyMatch=(a,b)=>{
    const na=normName(a),nb=normName(b);
    if(!na||!nb)return false;
    if(na===nb)return true;
    if(na.includes(nb)||nb.includes(na))return true;
    const pa=na.split(" "),pb=nb.split(" ");
    if(pa.length>=2&&pb.length>=2&&pa[0]===pb[0]&&pa[pa.length-1]===pb[pb.length-1])return true;
    return false;
  };
  // Find canonical name in map
  const findInMap=(map,name)=>{
    if(map[name])return name;
    for(const key of Object.keys(map)){if(fuzzyMatch(key,name))return key;}
    return null;
  };

  // Split multi-presenter string into individual names
  const splitPresenters=(str)=>{
    if(!str)return[];
    return str.split(/[,،؛;\/&]+|\s+و\s+|\s+and\s+/i).map(n=>n.trim()).filter(n=>n.length>1);
  };

  // Build staff data: split presenters + track attendance
  const staffMap=useMemo(()=>{
    const map={};
    const addToMap=(rawName,college)=>{
      if(!rawName||!rawName.trim())return null;
      const name=resolveCanonical(rawName);
      const existing=findInMap(map,name);
      if(existing)return existing;
      map[name]={name,college:college||"",title:"",email:"",phone:"",fromRoster:false,activities:[],jointActivities:[],attendedActivities:[],letters:[]};
      return name;
    };
    // Roster first
    (staffRoster||[]).forEach(r=>{
      const cn=resolveCanonical(r.name);
      if(map[cn]){
        // merge roster data into existing canonical
        map[cn].fromRoster=true;
        if(!map[cn].college&&r.college)map[cn].college=r.college;
        if(!map[cn].title&&r.title)map[cn].title=r.title;
        if(!map[cn].email&&r.email)map[cn].email=r.email;
        if(!map[cn].phone&&r.phone)map[cn].phone=r.phone;
      }else{
        map[cn]={name:cn,college:r.college||"",title:r.title||"",email:r.email||"",phone:r.phone||"",fromRoster:true,activities:[],jointActivities:[],attendedActivities:[],letters:[]};
      }
    });
    // Activities — split multi-presenter
    allActs.forEach(a=>{
      if(!a.presenter)return;
      const names=splitPresenters(a.presenter);
      const isJoint=names.length>1;
      names.forEach(name=>{
        const key=addToMap(name,a.college);
        if(!key)return;
        if(!map[key].activities.find(x=>x.id===a.id)){
          map[key].activities.push(a);
          if(isJoint)map[key].jointActivities.push(a);
        }
        if(a.college&&!map[key].college)map[key].college=a.college;
      });
    });
    // Attendance from activity participants field (plans + initiatives)
    allActs.forEach(a=>{
      if(!a.participants)return;
      const names=a.participants.split(/[,،\n;]+/).map(n=>n.trim()).filter(n=>n.length>1);
      names.forEach(name=>{
        // College-scoped fuzzy match: exact match anywhere, but fuzzy only within same college
        let key=null;
        const cName=resolveCanonical(name);
        if(map[cName]){key=cName;}
        else if(map[name]){key=name;}
        else{
          for(const k of Object.keys(map)){
            if(normName(k)===normName(name)){key=k;break;}
            if(a.college&&map[k].college===a.college&&fuzzyMatch(k,name)){key=k;break;}
          }
        }
        if(key){
          if(!map[key].attendedActivities.find(x=>x.id===a.id||x.title===a.title))map[key].attendedActivities.push(a);
        }
      });
    });
    // Attendance from event reports' participantNames
    (reports||[]).forEach(r=>{
      if(!r.participantNames)return;
      const attendees=r.participantNames.split(/[,،\n;]+/).map(n=>n.trim()).filter(n=>n.length>1);
      const act=allActs.find(a=>a.title===r.activityTitle)||{title:r.activityTitle,date:r.activityDate,type:r.activityType,college:r.college,status:"Completed"};
      attendees.forEach(name=>{
        // College-scoped fuzzy match
        let key=null;
        const cName=resolveCanonical(name);
        if(map[cName]){key=cName;}
        else if(map[name]){key=name;}
        else{
          for(const k of Object.keys(map)){
            if(normName(k)===normName(name)){key=k;break;}
            if(r.college&&map[k].college===r.college&&fuzzyMatch(k,name)){key=k;break;}
          }
        }
        if(key){
          if(!map[key].attendedActivities.find(x=>x.title===act.title))map[key].attendedActivities.push(act);
        }
        // Don't auto-create staff from attendance — only track for existing staff/roster
      });
    });
    // Letters
    (trainers||[]).forEach(tr=>{const key=findInMap(map,tr.name);if(key&&tr.appreciationLetters)map[key].letters=tr.appreciationLetters||[];});
    return map;
  },[allActs,trainers,staffRoster,reports,aliasMap]);

  const corrLetters=useMemo(()=>{
    const found={};
    (corr||[]).forEach(c=>{
      if(c.subject&&(c.subject.includes("شكر")||c.subject.toLowerCase().includes("appreciation")||c.subject.toLowerCase().includes("thank"))){
        Object.keys(staffMap).forEach(name=>{
          if(c.subject.includes(name)||c.notes?.includes(name)||(c.to||"").includes(name)){
            if(!found[name])found[name]=[];
            found[name].push({date:c.date,ref:c.refNumber,subject:c.subject,source:"correspondence"});
          }
        });
      }
    });
    return found;
  },[corr,staffMap]);

  const getStats=(s)=>{
    const completed=s.activities.filter(a=>a.status==="Completed").length;
    const overdue=s.activities.filter(a=>a.date&&new Date(a.date)<now&&a.status!=="Completed"&&a.status!=="Cancelled"&&a.status!=="Postponed").length;
    const cancelled=s.activities.filter(a=>a.status==="Cancelled").length;
    const planned=s.activities.length-completed-cancelled;
    const trData=trainers.find(tr=>fuzzyMatch(tr.name,s.name));
    const allLetters=[...(trData?.appreciationLetters||[]),...(corrLetters[s.name]||[])];
    const letters=[...new Map(allLetters.map(l=>[l.date+l.ref,l])).values()];
    const attended=s.attendedActivities.length;
    const joint=s.jointActivities.length;
    return{completed,overdue,cancelled,planned,letters,attendees:s.activities.reduce((sum,a)=>sum+(a.attendees||0),0),total:s.activities.length,attended,joint};
  };

  const staffList=useMemo(()=>{
    return Object.values(staffMap).filter(s=>{
      if(fCollege!=="All"&&s.college!==fCollege)return false;
      if(fStatus==="active"&&s.activities.length===0)return false;
      if(fStatus==="participated"&&s.attendedActivities.length===0)return false;
      if(fStatus==="inactive"&&s.activities.length>0)return false;
      if(fStatus==="noLetter"){const st=getStats(s);if(st.letters.length>0||s.activities.length===0)return false;}
      if(fSearch){const q=fSearch.toLowerCase();if(!s.name.toLowerCase().includes(q)&&!s.college.toLowerCase().includes(q)&&!(s.title||"").toLowerCase().includes(q))return false;}
      return true;
    }).sort((a,b)=>b.activities.length-a.activities.length);
  },[staffMap,fCollege,fSearch,fStatus]);

  const saveLetter=(staffName,letter)=>{
    const existing=trainers.find(tr=>tr.name===staffName);
    if(existing){saveTrainers(trainers.map(tr=>tr.name===staffName?{...tr,appreciationLetters:[...(tr.appreciationLetters||[]),letter]}:tr));}
    else{const s=staffMap[staffName];saveTrainers([...trainers,{id:gid(),name:staffName,college:s?.college||"",workplace:s?.college||"",degree:"",generalSpec:"",specificSpec:"",internal:true,notes:"",appreciationLetters:[letter],customFields:[]}]);}
    setLetterForm(null);
  };

  const addStaffMember=(d)=>{if(!d.name.trim())return;saveRoster([...(staffRoster||[]),{id:gid(),...d}]);setShowAddForm(false);};
  const removeFromRoster=(name)=>saveRoster((staffRoster||[]).filter(r=>r.name!==name));

  // Merge two staff into one canonical identity (persisted across sessions)
  const mergeStaff=(canonical,aliasName)=>{
    if(!canonical||!aliasName||canonical===aliasName)return;
    const list=[...(staffMerges||[])];
    // If alias was itself a canonical for other names, fold its aliases into the new canonical
    const aliasGroupIdx=list.findIndex(g=>g.canonical===aliasName);
    let foldedAliases=[];
    if(aliasGroupIdx>=0){foldedAliases=list[aliasGroupIdx].aliases||[];list.splice(aliasGroupIdx,1);}
    let group=list.find(g=>g.canonical===canonical);
    if(!group){group={canonical,aliases:[]};list.push(group);}
    const set=new Set([...(group.aliases||[]),aliasName,...foldedAliases]);
    set.delete(canonical);
    group.aliases=[...set];
    // Also fix any other group whose canonical is now an alias
    list.forEach(g=>{if(g!==group&&g.aliases?.includes(canonical)){g.aliases=g.aliases.filter(a=>a!==canonical);}});
    saveStaffMerges(list);
    // Rename roster entry from aliasName to canonical (so display matches)
    const updRoster=(staffRoster||[]).map(r=>r.name===aliasName?{...r,name:canonical}:r);
    saveRoster(updRoster);
    setMergeFor(null);
  };

  const unmergeStaff=(canonical,aliasName)=>{
    const list=(staffMerges||[]).map(g=>g.canonical===canonical?{...g,aliases:(g.aliases||[]).filter(a=>a!==aliasName)}:g).filter(g=>(g.aliases||[]).length>0);
    saveStaffMerges(list);
  };

  // Fuzzy name matching — handles partial names, titles, degrees
  const normalizeName=(name)=>{
    if(!name)return"";
    let n=name.trim();
    // Remove common titles/degrees
    n=n.replace(/^(د\.|أ\.د\.|أ\.م\.|م\.|م\.د\.|م\.م\.|Prof\.|Dr\.|Asst\.\s*Prof\.|Assoc\.\s*Prof\.|Mr\.|Mrs\.|Ms\.|Eng\.|Lect\.)\s*/gi,"");
    // Remove extra spaces
    n=n.replace(/\s+/g," ").trim().toLowerCase();
    return n;
  };

  const namesMatch=(a,b)=>{
    const na=normalizeName(a),nb=normalizeName(b);
    if(!na||!nb)return false;
    if(na===nb)return true;
    // Check if one contains the other (partial name)
    if(na.includes(nb)||nb.includes(na))return true;
    // Check first+last match (skip middle names)
    const pa=na.split(" "),pb=nb.split(" ");
    if(pa.length>=2&&pb.length>=2&&pa[0]===pb[0]&&pa[pa.length-1]===pb[pb.length-1])return true;
    return false;
  };

  const findDuplicate=(name,existingNames)=>{
    for(const existing of existingNames){
      if(namesMatch(name,existing))return existing;
    }
    return null;
  };

  const [importPreview,setImportPreview]=useState(null);

  const importStaff=async(file)=>{
    if(!file)return;setImportMsg("");
    try{
      const fname=file.name.toLowerCase();
      let rows=[];

      if(fname.endsWith(".json")){
        const text=await file.text();let data=JSON.parse(text);
        if(data.staff)data=data.staff;if(data.roster)data=data.roster;
        if(Array.isArray(data))rows=data;
        else{setImportMsg("⚠️ Invalid JSON format");return;}
      } else if(fname.endsWith(".xlsx")||fname.endsWith(".xls")){
        const buf=await file.arrayBuffer();
        const wb=XLSX.read(buf,{type:"array"});
        const ws=wb.Sheets[wb.SheetNames[0]];
        const raw=XLSX.utils.sheet_to_json(ws,{defval:""});
        // Auto-detect column names (flexible: name/الاسم, college/الكلية, etc.)
        rows=raw.map(r=>{
          const keys=Object.keys(r);
          const findCol=(patterns)=>keys.find(k=>patterns.some(p=>k.toLowerCase().includes(p)))||"";
          const nameCol=findCol(["name","اسم","الاسم","staff","الكادر","lecturer","محاضر"]);
          const collegeCol=findCol(["college","كلية","الكلية","faculty","department","قسم"]);
          const titleCol=findCol(["title","لقب","position","منصب","degree","شهادة","درجة"]);
          const emailCol=findCol(["email","بريد","الإيميل"]);
          const phoneCol=findCol(["phone","هاتف","موبايل","رقم"]);
          return{name:r[nameCol]||"",college:r[collegeCol]||"",title:r[titleCol]||"",email:r[emailCol]||"",phone:String(r[phoneCol]||"")};
        });
      } else if(fname.endsWith(".csv")||fname.endsWith(".txt")){
        const text=await file.text();
        const lines=text.split(/\r?\n/).map(l=>l.trim()).filter(Boolean);
        if(lines.length===0){setImportMsg("⚠️ Empty file");return;}
        // Detect if CSV (has commas/tabs) or plain list
        const sep=lines[0].includes("\t")?"\t":lines[0].includes(",")? ",":null;
        if(sep){
          // CSV with header
          const headers=lines[0].split(sep).map(h=>h.trim().replace(/^"|"$/g,""));
          const findH=(patterns)=>headers.findIndex(h=>patterns.some(p=>h.toLowerCase().includes(p)));
          const ni=findH(["name","اسم","الاسم","staff","lecturer"]);
          const ci=findH(["college","كلية","الكلية","faculty","department"]);
          const ti=findH(["title","لقب","degree","position"]);
          for(let i=1;i<lines.length;i++){
            const cols=lines[i].split(sep).map(c=>c.trim().replace(/^"|"$/g,""));
            if(cols[ni>=0?ni:0])rows.push({name:cols[ni>=0?ni:0],college:ci>=0?cols[ci]||"":"",title:ti>=0?cols[ti]||"":""});
          }
        }else{
          // Plain text: one name per line, optionally "Name - College"
          let currentCollege="";
          lines.forEach(line=>{
            // Check if this line is a college header (starts with "College of" or "كلية")
            if(line.match(/^(College of|Department of|كلية|قسم)\s/i)){currentCollege=line;return;}
            if(line.startsWith("#")||line.startsWith("//")){return;}// skip comments
            const parts=line.split(/\s*[-–—|]\s*/);
            const name=parts[0].trim();
            const college=parts[1]?.trim()||currentCollege;
            if(name)rows.push({name,college,title:parts[2]?.trim()||""});
          });
        }
      }else{setImportMsg("⚠️ Supported: .xlsx, .csv, .txt, .json");return;}

      // Filter valid rows
      rows=rows.filter(r=>r.name&&r.name.trim());
      if(rows.length===0){setImportMsg("⚠️ No names found in file");return;}

      // Fuzzy duplicate check against existing roster + activity presenters
      const existingNames=[...(staffRoster||[]).map(r=>r.name),...Object.keys(staffMap)];
      const results=rows.map(r=>{
        const match=findDuplicate(r.name,existingNames);
        return{...r,id:gid(),duplicate:!!match,matchedWith:match||""};
      });

      const dupes=results.filter(r=>r.duplicate).length;
      if(dupes>0||results.length>10){
        // Show preview for review
        setImportPreview(results);
        setImportMsg(`📋 ${results.length} found, ${dupes} possible duplicates — review below`);
      }else{
        // Auto-import (small list, no dupes)
        const clean=results.map(({duplicate,matchedWith,...r})=>r);
        saveRoster([...(staffRoster||[]),...clean]);
        setImportMsg(`✅ ${clean.length} staff imported`);
        setTimeout(()=>setImportMsg(""),3000);
      }
    }catch(e){setImportMsg("❌ "+e.message);setTimeout(()=>setImportMsg(""),4000);}
  };

  const confirmImport=(items)=>{
    const clean=items.map(({duplicate,matchedWith,selected,...r})=>r);
    saveRoster([...(staffRoster||[]),...clean]);
    setImportPreview(null);setImportMsg(`✅ ${clean.length} staff imported`);
    setTimeout(()=>setImportMsg(""),3000);
  };

  const totalStaff=Object.keys(staffMap).length;const activeCount=Object.values(staffMap).filter(s=>s.activities.length>0).length;const inactiveCount=totalStaff-activeCount;

  const exportStaffExcel=()=>{
    const wb=XLSX.utils.book_new();
    const data=staffList.map(s=>{const st=getStats(s);return{"Name":s.name,"College":s.college||"","Title":s.title||"","Presented":st.total,"Completed":st.completed,"Overdue":st.overdue,"Cancelled":st.cancelled,"Planned":st.planned,"Attended":st.attended,"Joint":st.joint,"Attendees":st.attendees,"Letters":st.letters.length,"Source":s.fromRoster?"Roster":"Activity"};});
    XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(data),"Staff Summary");
    // Letters sheet
    const letters=[];staffList.forEach(s=>{const st=getStats(s);st.letters.forEach(l=>{letters.push({"Staff":s.name,"College":s.college||"","Date":l.date||"","Reference":l.ref||"","Source":l.source||"manual","Notes":l.notes||""});});});
    if(letters.length>0)XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(letters),"Appreciation Letters");
    // Inactive
    const inactive=staffList.filter(s=>s.activities.length===0);
    if(inactive.length>0)XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(inactive.map(s=>({"Name":s.name,"College":s.college||"","Title":s.title||"","Source":s.fromRoster?"Roster":"Activity"}))),"No Activity");
    XLSX.writeFile(wb,`Staff_${fCollege!=="All"?fCollege.replace(/College of |Department of /g,""):"All"}_${new Date().toISOString().split('T')[0]}.xlsx`);
  };

  const printStaff=()=>{
    const isRtl=t.dir==="rtl";const al=isRtl?"right":"left";
    let html=`<html dir="${isRtl?'rtl':'ltr'}"><head><meta charset="utf-8"><style>@import url('https://fonts.googleapis.com/css2?family=Noto+Sans+Arabic:wght@400;700&display=swap');body{font-family:'Noto Sans Arabic','Noto Sans',sans-serif;padding:30px;font-size:12px;}h1{color:#1a5276;font-size:20px;}h2{color:#1a5276;font-size:15px;border-bottom:2px solid #1a5276;padding-bottom:4px;}table{width:100%;border-collapse:collapse;margin:8px 0 16px;}th,td{border:1px solid #ddd;padding:5px 8px;text-align:${al};font-size:11px;}th{background:#f4f6f7;font-weight:700;}.badge{padding:2px 8px;border-radius:8px;font-size:10px;font-weight:600;}.stat{display:inline-block;padding:6px 14px;margin:0 6px 6px 0;border-radius:8px;background:#f4f6f7;text-align:center;}.stat b{display:block;font-size:18px;color:#1a5276;}</style></head><body>`;
    if(logoSrc)html+=`<div style="display:flex;align-items:center;gap:14px;margin-bottom:16px;"><img src="${logoSrc}" width="55" style="border-radius:50%;background:#0d3249;"/><div><h1 style="margin:0">${pName||""}</h1><p style="margin:2px 0;font-size:12px;color:#6c7a89;">${t.staffPage} — ${fCollege!=="All"?fCollege:"All Colleges"} — ${new Date().toLocaleDateString()}</p></div></div>`;
    html+=`<div style="margin-bottom:14px;"><div class="stat"><b>${totalStaff}</b>Total</div><div class="stat"><b style="color:#27ae60">${activeCount}</b>Active</div><div class="stat"><b style="color:#e74c3c">${inactiveCount}</b>No activity</div></div>`;
    html+=`<h2>${t.staffPage} (${staffList.length})</h2><table><tr><th>#</th><th>Name</th><th>College</th><th>Presented</th><th>Completed</th><th>Overdue</th><th>Attended</th><th>Joint</th><th>Letters</th></tr>`;
    staffList.forEach((s,i)=>{const st=getStats(s);html+=`<tr><td>${i+1}</td><td>${s.name}${s.title?" · "+s.title:""}</td><td>${(s.college||"").replace("College of ","").replace("Department of ","")}</td><td>${st.total}</td><td style="color:#27ae60;font-weight:600">${st.completed}</td><td style="color:${st.overdue>0?"#e74c3c":"#27ae60"};font-weight:${st.overdue>0?700:400}">${st.overdue}</td><td>${st.attended}</td><td>${st.joint}</td><td>${st.letters.length}</td></tr>`;});
    html+=`</table>`;
    const inactiveList=staffList.filter(s=>s.activities.length===0);
    if(inactiveList.length>0){html+=`<h2>⚠️ No Activity (${inactiveList.length})</h2><table><tr><th>#</th><th>Name</th><th>College</th><th>Title</th></tr>`;inactiveList.forEach((s,i)=>{html+=`<tr><td>${i+1}</td><td>${s.name}</td><td>${(s.college||"").replace("College of ","").replace("Department of ","")}</td><td>${s.title||""}</td></tr>`;});html+=`</table>`;}
    html+=`</body></html>`;showPrintPreview(html);
  };

  return(
    <div>
      <div style={{display:"flex",justifyContent:"space-between",marginBottom:12,flexWrap:"wrap",gap:8}}>
        <div><h2 style={{fontSize:18,fontWeight:700,margin:0}}>{t.staffPage}</h2><p style={{fontSize:12,color:"var(--color-text-secondary)",margin:"3px 0 0"}}>{t.staffSubtitle}</p></div>
        {isAdmin&&<div style={{display:"flex",gap:6,flexWrap:"wrap"}}>
          <button onClick={()=>setShowAddForm(!showAddForm)} style={{...btnP,padding:"6px 14px",fontSize:11,display:"flex",alignItems:"center",gap:4}}><Ic n="plus" s={13} c="#fff"/>{t.add}</button>
          <label style={{...btnP,padding:"6px 14px",fontSize:11,display:"inline-flex",alignItems:"center",gap:4,cursor:"pointer",background:"#2980b9"}}>📥 Import<input type="file" accept=".xlsx,.xls,.csv,.txt,.json" onChange={e=>{importStaff(e.target.files[0]);e.target.value="";}} style={{display:"none"}}/></label>
          <button onClick={exportStaffExcel} style={{...btnP,padding:"6px 14px",fontSize:11,background:"#217346",display:"flex",alignItems:"center",gap:4}}>📊 Excel</button>
          <button onClick={printStaff} style={{...btnP,padding:"6px 14px",fontSize:11,background:"#c0392b",display:"flex",alignItems:"center",gap:4}}>🖨️ Print</button>
        </div>}
      </div>

      {/* SUMMARY */}
      <div style={{display:"flex",gap:8,marginBottom:12,flexWrap:"wrap"}}>
        <span style={{fontSize:11,padding:"4px 10px",borderRadius:8,background:"#eaf2f8",color:"#1a5276",fontWeight:600}}>👥 {totalStaff} total</span>
        <span style={{fontSize:11,padding:"4px 10px",borderRadius:8,background:"#d5f5e3",color:"#196f3d",fontWeight:600}}>✅ {activeCount} active</span>
        <span style={{fontSize:11,padding:"4px 10px",borderRadius:8,background:inactiveCount>0?"#fdedec":"#f5f5f5",color:inactiveCount>0?"#c0392b":"#888",fontWeight:600}}>⚠️ {inactiveCount} no activity</span>
        {(staffRoster||[]).length>0&&<span style={{fontSize:11,padding:"4px 10px",borderRadius:8,background:"#f0e6f6",color:"#6c3483"}}>📋 {(staffRoster||[]).length} from roster</span>}
      </div>

      {/* FILTERS */}
      <div style={{display:"flex",gap:8,marginBottom:14,flexWrap:"wrap",alignItems:"center"}}>
        <select value={fCollege} onChange={e=>setFCollege(e.target.value)} style={{...inp,width:"auto",minWidth:160,padding:"5px 10px",fontSize:11}}>
          <option value="All">🏛️ {t.all} {t.college}</option>
          {(COLLEGES||[]).map(c=><option key={c}>{c}</option>)}
        </select>
        <select value={fStatus} onChange={e=>setFStatus(e.target.value)} style={{...inp,width:"auto",minWidth:130,padding:"5px 10px",fontSize:11}}>
          <option value="All">{t.all}</option>
          <option value="active">✅ Active (has presentations)</option>
          <option value="participated">🎓 Has attendance records</option>
          <option value="inactive">⚠️ No presentations</option>
          <option value="noLetter">📜 Active, no letter</option>
        </select>
        <input value={fSearch} onChange={e=>setFSearch(e.target.value)} placeholder={"🔍 "+t.search+"..."} style={{...inp,width:"auto",minWidth:180,padding:"5px 10px",fontSize:11}}/>
        <span style={{fontSize:11,color:"var(--color-text-tertiary)"}}>{staffList.length}/{totalStaff}</span>
      </div>
      {importMsg&&<div style={{fontSize:12,marginBottom:10,color:importMsg.includes("✅")?"#27ae60":importMsg.includes("⚠️")||importMsg.includes("📋")?"#e67e22":"#e74c3c"}}>{importMsg}</div>}

      {/* IMPORT PREVIEW */}
      {importPreview&&<div style={{...crd,border:"2px solid #2980b9",marginBottom:14,maxHeight:400,overflow:"auto"}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:10}}>
          <h3 style={{fontSize:13,fontWeight:600,margin:0}}>📋 Review import ({importPreview.length} staff)</h3>
          <div style={{display:"flex",gap:6}}>
            <button onClick={()=>{const noDupes=importPreview.filter(r=>!r.duplicate);confirmImport(noDupes);}} style={{...btnP,padding:"5px 12px",fontSize:11}}>Import {importPreview.filter(r=>!r.duplicate).length} new</button>
            <button onClick={()=>{confirmImport(importPreview);}} style={{...btnP,padding:"5px 12px",fontSize:11,background:"#e67e22"}}>Import all {importPreview.length}</button>
            <button onClick={()=>{setImportPreview(null);setImportMsg("");}} style={{...btnC,padding:"5px 12px",fontSize:11}}>Cancel</button>
          </div>
        </div>
        <div style={{fontSize:10,color:"var(--color-text-tertiary)",marginBottom:8}}>🟡 = possible duplicate (fuzzy name match). Review and decide what to import.</div>
        {importPreview.map((r,i)=>(
          <div key={i} style={{display:"flex",alignItems:"center",gap:8,padding:"5px 8px",borderBottom:"0.5px solid var(--color-border-tertiary)",fontSize:11,background:r.duplicate?"#fef9e7":"transparent"}}>
            <span style={{width:20,textAlign:"center",color:"var(--color-text-tertiary)"}}>{i+1}</span>
            <span style={{flex:1,fontWeight:500}}>{r.name}</span>
            <span style={{fontSize:10,color:"var(--color-text-tertiary)",minWidth:100}}>{r.college?r.college.replace("College of ","").replace("Department of ",""):""}</span>
            {r.title&&<span style={{fontSize:10,color:"var(--color-text-tertiary)"}}>{r.title}</span>}
            {r.duplicate&&<span style={{fontSize:9,padding:"2px 6px",borderRadius:6,background:"#f39c12",color:"#fff",whiteSpace:"nowrap"}}>🟡 ≈ {r.matchedWith}</span>}
            {!r.duplicate&&<span style={{fontSize:9,padding:"2px 6px",borderRadius:6,background:"#27ae60",color:"#fff"}}>✓ new</span>}
            <button onClick={()=>setImportPreview(importPreview.filter((_,idx)=>idx!==i))} style={{background:"none",border:"none",cursor:"pointer",color:"#e74c3c",fontSize:11,padding:0}}>✕</button>
          </div>
        ))}
      </div>}

      {/* ADD FORM */}
      {showAddForm&&<div style={{...crd,border:"1.5px solid #c9a84c",marginBottom:14}}>
        <h3 style={{fontSize:13,fontWeight:600,margin:"0 0 8px"}}>{t.add} {t.staffMember}</h3>
        <StaffAddForm onSave={addStaffMember} onCancel={()=>setShowAddForm(false)} t={t} COLLEGES={COLLEGES}/>
      </div>}

      {/* STAFF CARDS */}
      {staffList.map(s=>{
        const st=getStats(s);const inactive=s.activities.length===0;
        return(
          <div key={s.name} style={{...crd,marginBottom:10,padding:"14px 18px",border:inactive?"1.5px dashed #e74c3c":"0.5px solid var(--color-border-tertiary)",opacity:inactive?0.85:1}}>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",flexWrap:"wrap",gap:8}}>
              <div style={{flex:1,minWidth:200}}>
                <div style={{display:"flex",gap:8,alignItems:"center",marginBottom:6}}>
                  <div style={{width:36,height:36,borderRadius:"50%",background:inactive?"#fdedec":"#0d324915",display:"flex",alignItems:"center",justifyContent:"center",fontSize:14,fontWeight:600,color:inactive?"#e74c3c":"#1a5276"}}>{s.name[0]}</div>
                  <div style={{flex:1}}>
                    <div style={{fontSize:14,fontWeight:600}}>{s.name}{s.title&&<span style={{fontSize:11,color:"var(--color-text-tertiary)",fontWeight:400}}> · {s.title}</span>}</div>
                    <div style={{fontSize:11,color:"var(--color-text-tertiary)"}}>
                      {s.college?s.college.replace("College of ","").replace("Department of ",""):""}{s.fromRoster&&" 📋"}{inactive?" · ⚠️ No activities":" · "+s.activities.length+" "+t.acts}
                    </div>
                  </div>
                  {isAdmin&&<button onClick={()=>setMergeFor(s.name)} style={{...btnI,padding:"2px 6px",fontSize:11,color:"#2980b9"}} title="Merge into another staff">🔗</button>}
                  {isAdmin&&s.fromRoster&&<button onClick={()=>removeFromRoster(s.name)} style={{...btnI,padding:2}} title="Remove"><Ic n="trash" s={12} c="#e74c3c"/></button>}
                </div>
                {!inactive&&<>
                  <div style={{display:"flex",gap:6,flexWrap:"wrap",marginBottom:6}}>
                    <span style={{fontSize:10,padding:"2px 8px",borderRadius:8,background:"#d5f5e3",color:"#196f3d",fontWeight:600}}>✅ {st.completed}</span>
                    {st.overdue>0&&<span style={{fontSize:10,padding:"2px 8px",borderRadius:8,background:"#fdedec",color:"#c0392b",fontWeight:600}}>⚠️ {st.overdue}</span>}
                    {st.cancelled>0&&<span style={{fontSize:10,padding:"2px 8px",borderRadius:8,background:"#f5f5f5",color:"#888",fontWeight:600}}>❌ {st.cancelled}</span>}
                    <span style={{fontSize:10,padding:"2px 8px",borderRadius:8,background:"#eaf2f8",color:"#1a5276"}}>📋 {st.planned} planned</span>
                    {st.attendees>0&&<span style={{fontSize:10,padding:"2px 8px",borderRadius:8,background:"#f0e6f6",color:"#6c3483"}}>👥 {st.attendees}</span>}
                    {st.attended>0&&<span style={{fontSize:10,padding:"2px 8px",borderRadius:8,background:"#e8f8f5",color:"#117864"}}>🎓 {st.attended} attended</span>}
                    {st.joint>0&&<span style={{fontSize:10,padding:"2px 8px",borderRadius:8,background:"#fef9e7",color:"#7d6608"}}>🤝 {st.joint} joint</span>}
                  </div>
                  <div style={{fontSize:11,color:"var(--color-text-secondary)"}}>
                    {s.activities.slice(0,5).map(a=>(<div key={a.id} style={{display:"flex",alignItems:"center",gap:6,padding:"3px 0",cursor:"pointer"}} onClick={()=>setViewAct(a)}><div style={{width:6,height:6,borderRadius:"50%",background:sC(a.status),flexShrink:0}}/><span style={{flex:1,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{a.title}</span><span style={{fontSize:9,color:"var(--color-text-tertiary)",flexShrink:0}}>{a.date||""}</span></div>))}
                    {s.activities.length>5&&<div style={{fontSize:10,color:"var(--color-text-tertiary)",paddingTop:3}}>...+{s.activities.length-5} more</div>}
                  </div>
                </>}
              </div>
              <div style={{minWidth:170,padding:"8px 12px",background:"#fef9e7",borderRadius:8,border:"1px solid #f0e6a0"}}>
                <div style={{fontSize:11,fontWeight:600,color:"#7d6608",marginBottom:4}}>📜 {t.appreciationLetter} ({st.letters.length})</div>
                {st.letters.map((l,i)=>(<div key={i} style={{fontSize:10,color:"#5d4e0e",padding:"2px 0",borderBottom:"0.5px solid #f0e6a0"}}>{l.date} {l.ref&&`· ${l.ref}`} {l.source==="correspondence"&&"📨"}</div>))}
                {st.letters.length===0&&<div style={{fontSize:10,color:"#b7950b",fontStyle:"italic"}}>—</div>}
                {isAdmin&&<button onClick={()=>setLetterForm(s.name)} style={{...btnC,padding:"3px 8px",fontSize:10,marginTop:4}}>{t.addLetter}</button>}
              </div>
            </div>
          </div>
        );
      })}
      {staffList.length===0&&<div style={{textAlign:"center",padding:50,color:"var(--color-text-tertiary)"}}>{t.noStaffData}</div>}

      {/* ACTIVITY POPUP */}
      {viewAct&&<div onClick={()=>setViewAct(null)} style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.45)",display:"flex",alignItems:"center",justifyContent:"center",zIndex:999}}>
        <div onClick={e=>e.stopPropagation()} style={{background:"#ffffff",borderRadius:14,padding:24,maxWidth:500,width:"92%",maxHeight:"80vh",overflow:"auto"}}>
          <h3 style={{fontSize:16,fontWeight:700,margin:"0 0 12px",color:"#1a5276"}}>{viewAct.title}</h3>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8,fontSize:12}}>
            {[{l:t.type,v:viewAct.type},{l:t.date,v:viewAct.date},{l:t.status,v:viewAct.status},{l:t.source,v:viewAct.source},{l:t.college,v:viewAct.college},{l:t.presenter,v:viewAct.presenter},{l:t.delivery,v:viewAct.delivery},{l:t.target,v:viewAct.target},{l:t.venue,v:viewAct.venue},{l:t.duration,v:viewAct.duration},{l:t.attendees,v:viewAct.attendees}].filter(x=>x.v).map((x,i)=>(<div key={i}><span style={{fontSize:10,color:"var(--color-text-tertiary)"}}>{x.l}</span><div style={{fontWeight:500}}>{x.v}</div></div>))}
          </div>
          {viewAct.objectives&&<div style={{marginTop:10,fontSize:12}}><span style={{fontWeight:600}}>{t.objectives}:</span> {viewAct.objectives}</div>}
          {viewAct.notes&&<div style={{marginTop:6,fontSize:12,color:"var(--color-text-secondary)"}}>{viewAct.notes}</div>}
          <button onClick={()=>setViewAct(null)} style={{...btnP,marginTop:14}}>{t.cancel}</button>
        </div>
      </div>}

      {/* MERGE MODAL */}
      {mergeFor&&<div onClick={()=>setMergeFor(null)} style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.45)",display:"flex",alignItems:"center",justifyContent:"center",zIndex:999}}>
        <div onClick={e=>e.stopPropagation()} style={{background:"#ffffff",borderRadius:14,padding:24,maxWidth:520,width:"92%",maxHeight:"80vh",overflow:"auto"}}>
          <h3 style={{fontSize:14,fontWeight:600,margin:"0 0 8px"}}>🔗 Merge staff</h3>
          <div style={{fontSize:12,color:"#555",marginBottom:10}}>Treat <b>{mergeFor}</b> as the same person as another staff member. All activities, attendance and letters will be combined under the target's name. This is remembered for future imports.</div>
          <MergePicker current={mergeFor} options={Object.keys(staffMap).filter(n=>n!==mergeFor).sort()} onPick={target=>mergeStaff(target,mergeFor)} onCancel={()=>setMergeFor(null)}/>
          {(staffMerges||[]).length>0&&<div style={{marginTop:14,paddingTop:10,borderTop:"1px solid #eee"}}>
            <div style={{fontSize:11,fontWeight:600,marginBottom:6,color:"#555"}}>Existing merges</div>
            {(staffMerges||[]).map((g,i)=>(
              <div key={i} style={{fontSize:11,padding:"4px 0",borderBottom:"0.5px solid #f0f0f0"}}>
                <b>{g.canonical}</b> ← {(g.aliases||[]).map((a,j)=>(<span key={j} style={{display:"inline-flex",alignItems:"center",gap:3,marginRight:6}}>{a}<button onClick={()=>unmergeStaff(g.canonical,a)} style={{background:"none",border:"none",color:"#e74c3c",cursor:"pointer",padding:0,fontSize:11}} title="Unmerge">✕</button></span>))}
              </div>
            ))}
          </div>}
        </div>
      </div>}

      {/* LETTER FORM */}
      {letterForm&&<div onClick={()=>setLetterForm(null)} style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.45)",display:"flex",alignItems:"center",justifyContent:"center",zIndex:999}}>
        <div onClick={e=>e.stopPropagation()} style={{background:"#ffffff",borderRadius:14,padding:24,maxWidth:400,width:"92%"}}>
          <h3 style={{fontSize:14,fontWeight:600,margin:"0 0 12px"}}>📜 {t.addLetter} — {letterForm}</h3>
          <LetterForm onSave={l=>saveLetter(letterForm,l)} onCancel={()=>setLetterForm(null)} t={t}/>
        </div>
      </div>}
    </div>
  );
}

function MergePicker({current,options,onPick,onCancel}){
  const [q,setQ]=useState("");
  const [sel,setSel]=useState("");
  const filtered=options.filter(n=>n.toLowerCase().includes(q.toLowerCase())).slice(0,200);
  return(<div>
    <input autoFocus value={q} onChange={e=>setQ(e.target.value)} placeholder="🔍 Search target staff..." style={{...inp,marginBottom:8}}/>
    <div style={{maxHeight:240,overflow:"auto",border:"1px solid #eee",borderRadius:6,marginBottom:10}}>
      {filtered.length===0&&<div style={{padding:10,fontSize:11,color:"#888"}}>No matches</div>}
      {filtered.map(n=>(
        <div key={n} onClick={()=>setSel(n)} style={{padding:"6px 10px",fontSize:12,cursor:"pointer",background:sel===n?"#eaf2f8":"transparent",borderBottom:"0.5px solid #f5f5f5"}}>{n}</div>
      ))}
    </div>
    <div style={{display:"flex",gap:8,justifyContent:"flex-end"}}>
      <button onClick={onCancel} style={btnC}>Cancel</button>
      <button onClick={()=>sel&&onPick(sel)} style={{...btnP,opacity:sel?1:0.5}}>Merge "{current}" → "{sel||"…"}"</button>
    </div>
  </div>);
}

function StaffAddForm({onSave,onCancel,t,COLLEGES}){
  const [d,sD]=useState({name:"",college:"",title:"",email:"",phone:""});
  return(<div>
    <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:12}}>
      <div><label style={lbl}>{t.name}</label><input value={d.name} onChange={e=>sD({...d,name:e.target.value})} style={inp}/></div>
      <div><label style={lbl}>{t.college}</label><select value={d.college} onChange={e=>sD({...d,college:e.target.value})} style={inp}><option value="">--</option>{(COLLEGES||[]).map(c=><option key={c}>{c}</option>)}</select></div>
      <div><label style={lbl}>Title</label><input value={d.title} onChange={e=>sD({...d,title:e.target.value})} style={inp} placeholder="e.g. Lecturer, Professor"/></div>
      <div><label style={lbl}>{t.email||"Email"}</label><input value={d.email} onChange={e=>sD({...d,email:e.target.value})} style={inp}/></div>
    </div>
    <div style={{display:"flex",gap:8,justifyContent:"flex-end"}}><button onClick={onCancel} style={btnC}>{t.cancel}</button><button onClick={()=>d.name.trim()&&onSave(d)} style={{...btnP,opacity:d.name.trim()?1:0.5}}>{t.save}</button></div>
  </div>);
}

function LetterForm({onSave,onCancel,t}){
  const [d,sD]=useState({date:new Date().toISOString().split("T")[0],ref:"",notes:""});
  return(<div>
    <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:12}}>
      <div><label style={lbl}>{t.letterDate}</label><input type="date" value={d.date} onChange={e=>sD({...d,date:e.target.value})} style={inp}/></div>
      <div><label style={lbl}>{t.letterRef}</label><input value={d.ref} onChange={e=>sD({...d,ref:e.target.value})} style={inp}/></div>
      <div style={{gridColumn:"1/-1"}}><label style={lbl}>{t.notes}</label><input value={d.notes} onChange={e=>sD({...d,notes:e.target.value})} style={inp}/></div>
    </div>
    <div style={{display:"flex",gap:8,justifyContent:"flex-end"}}><button onClick={onCancel} style={btnC}>{t.cancel}</button><button onClick={()=>d.date&&onSave(d)} style={{...btnP,opacity:d.date?1:0.5}}>{t.save}</button></div>
  </div>);
}

// ─── REPORTS ──────────────────────────────────────────────────────────
function ReportsPage({t,reports,saveReports,allActs,isAdmin,canEdit,logoSrc,pName,COLLEGES,plans,savePlans,inits,saveInits}){
  const [sf,sSf]=useState(false);
  const [editId,setEditId]=useState(null);
  const [fCollege,setFCollege]=useState("All");const [fSource,setFSource]=useState("All");const [fSearch,setFSearch]=useState("");
  const add=d=>{
    saveReports([...reports,{id:gid(),createdDate:new Date().toISOString().split("T")[0],...d}]);
    syncToActivity(d);
    sSf(false);
  };
  const del=id=>saveReports(reports.filter(r=>r.id!==id));

  const syncToActivity=(d)=>{
    if(!d.activityId)return;
    const actId=d.activityId;
    const updates={attendees:d.totalParticipants||0,participants:d.participantNames||"",status:"Completed"};
    // Sync all media fields from report to activity
    if(d.mediaStatus&&d.mediaStatus!=="Not contacted")updates.mediaStatus=d.mediaStatus;
    if(d.mediaChannel)updates.mediaChannel=d.mediaChannel;
    if(d.mediaContactedDate)updates.mediaContactedDate=d.mediaContactedDate;
    if(d.mediaPublishedDate)updates.mediaPublishedDate=d.mediaPublishedDate;
    if(d.mediaCoverageNotes)updates.mediaCoverageNotes=d.mediaCoverageNotes;
    if(d.mediaLinks)updates.mediaLinks=d.mediaLinks;
    let found=false;
    const updPlans=plans.map(p=>{
      const hasAct=p.activities.find(a=>a.id===actId);
      if(hasAct){found=true;return{...p,activities:p.activities.map(a=>a.id===actId?{...a,...updates}:a)};}
      return p;
    });
    if(found)savePlans(updPlans);
    else saveInits(inits.map(i=>i.id===actId?{...i,...updates}:i));
  };

  const upd=(d)=>{
    saveReports(reports.map(r=>r.id===editId?{...r,...d}:r));
    syncToActivity(d);
    setEditId(null);
  };
  const filteredReports=reports.filter(r=>{
    if(fCollege!=="All"&&r.college!==fCollege)return false;
    if(fSource!=="All"&&r.source!==fSource)return false;
    if(fSearch){const q=fSearch.toLowerCase();const s=[r.activityTitle,r.college,r.presenter,r.summary,r.participantNames,r.activityDate,r.activityType,r.source].filter(Boolean).join(" ").toLowerCase();if(!s.includes(q))return false;}
    return true;
  });
  const printReport=r=>{
    const useLogo=logoSrc||LOGO;
    const html=`<html dir="${t.dir}"><head><title>${r.activityTitle}</title><style>body{font-family:${t.font};padding:40px;} h1{color:#1a5276;} .f{display:flex;gap:20px;margin:8px 0;} .l{font-weight:bold;min-width:150px;}</style></head><body><img src="${useLogo}" width="60" style="border-radius:50%;margin-bottom:10px;"><h1>${r.activityTitle}</h1><div class="f"><span class="l">${t.date}:</span><span>${r.activityDate}</span></div>${r.college?`<div class="f"><span class="l">${t.college}:</span><span>${r.college}</span></div>`:""}<div class="f"><span class="l">${t.presenter}:</span><span>${r.presenter||"-"}</span></div><div class="f"><span class="l">${t.totalParticipants}:</span><span>${r.totalParticipants}</span></div><div class="f"><span class="l">${t.summary}:</span><span>${r.summary||"-"}</span></div><div class="f"><span class="l">${t.participantNames}:</span><span>${r.participantNames||"-"}</span></div><h2 style="font-size:14px;color:#e67e22;">📺 ${t.mediaCoverage}</h2><div class="f"><span class="l">${t.mediaStatus}:</span><span>${r.mediaStatus||"Not contacted"}</span></div>${r.mediaChannel?`<div class="f"><span class="l">${t.mediaChannel}:</span><span>${r.mediaChannel}</span></div>`:""}${r.mediaContactedDate?`<div class="f"><span class="l">${t.mediaContactedDate}:</span><span>${r.mediaContactedDate}</span></div>`:""}${r.mediaPublishedDate?`<div class="f"><span class="l">${t.mediaPublishedDate}:</span><span>${r.mediaPublishedDate}</span></div>`:""}${r.mediaLinks?`<div class="f"><span class="l">${t.mediaLinks}:</span><span>${r.mediaLinks}</span></div>`:""}${r.mediaCoverageNotes?`<div class="f"><span class="l">${t.mediaCoverageNotes}:</span><span>${r.mediaCoverageNotes}</span></div>`:""}<div class="f"><span class="l">${t.confirmed}:</span><span>${r.confirmed?"✅":"❌"}</span></div><br><p style="color:#888;font-size:12px;">Report filed: ${r.createdDate}</p></body></html>`;
    showPrintPreview(html);
  };

  return(
    <div>
      <div style={{display:"flex",justifyContent:"space-between",marginBottom:16,flexWrap:"wrap",gap:8}}>
        <div><h2 style={{fontSize:18,fontWeight:700,margin:0}}>{t.reports}</h2><p style={{fontSize:12,color:"var(--color-text-secondary)",margin:"3px 0 0"}}>{t.reportSubtitle}</p></div>
        <div style={{display:"flex",gap:8,alignItems:"center",flexWrap:"wrap"}}>
          <select value={fCollege} onChange={e=>setFCollege(e.target.value)} style={{...inp,width:"auto",minWidth:160,padding:"5px 10px",fontSize:11}}>
            <option value="All">🏛️ {t.all} {t.college}</option>
            {(COLLEGES||[]).map(c=><option key={c}>{c}</option>)}
          </select>
          <select value={fSource} onChange={e=>setFSource(e.target.value)} style={{...inp,width:"auto",minWidth:140,padding:"5px 10px",fontSize:11}}>
            <option value="All">{t.all} {t.source}</option>
            {SOURCES.map(s=><option key={s}>{s}</option>)}
          </select>
          <input value={fSearch} onChange={e=>setFSearch(e.target.value)} placeholder={"🔍 "+t.search+"..."} style={{...inp,width:"auto",minWidth:180,padding:"5px 10px",fontSize:11}}/>
          <span style={{fontSize:11,color:"var(--color-text-tertiary)"}}>{filteredReports.length}/{reports.length}</span>
          <button onClick={()=>exportReportsExcel(filteredReports,t)} style={{...btnP,padding:"5px 12px",fontSize:11,background:"#217346",display:"flex",alignItems:"center",gap:4}}>📊 {t.exportAsExcel}</button>
          <button onClick={()=>printReports(filteredReports,t,logoSrc,pName)} style={{...btnP,padding:"5px 12px",fontSize:11,background:"#c0392b",display:"flex",alignItems:"center",gap:4}}>🖨️ {t.exportAsPDF}</button>
          {(isAdmin||canEdit)&&<button onClick={()=>sSf(true)} style={{...btnP,display:"flex",alignItems:"center",gap:5}}><Ic n="plus" s={15} c="#fff"/>{t.newReport}</button>}
        </div>
      </div>
      {sf&&<RptForm t={t} allActs={allActs} onSave={add} onCancel={()=>sSf(false)} COLLEGES={COLLEGES}/>}
      {editId&&<RptForm t={t} allActs={allActs} onSave={upd} onCancel={()=>setEditId(null)} COLLEGES={COLLEGES} initial={reports.find(r=>r.id===editId)}/>}
      {filteredReports.map(r=>(
        <div key={r.id} style={{...crd,marginBottom:10,padding:"14px 18px"}}>
          <div style={{display:"flex",justifyContent:"space-between"}}>
            <div style={{flex:1,minWidth:0}}>
              <div style={{display:"flex",gap:6,alignItems:"center",flexWrap:"wrap"}}>
                <span style={{fontSize:14,fontWeight:600}}>{r.activityTitle}</span>
                {r.college&&<span style={{fontSize:9,padding:"2px 7px",borderRadius:8,background:"#e8d5f5",color:"#6c3483"}}>{r.college.replace("College of ","").replace("Department of ","")}</span>}
                {r.source&&<span style={{fontSize:9,padding:"2px 7px",borderRadius:8,background:r.source.includes("Initiative")||r.source.includes("Directive")?"#fef5e7":"#eaf2f8",color:r.source.includes("Initiative")||r.source.includes("Directive")?"#ca6f1e":"#2471a3"}}>{r.source}</span>}
                {r.activityType&&<span style={{fontSize:9,padding:"2px 7px",borderRadius:8,background:"var(--color-background-secondary)",color:"var(--color-text-secondary)"}}>{r.activityType}</span>}
              </div>
              <div style={{fontSize:11,color:"var(--color-text-secondary)",marginTop:3}}>{t.date}: {r.activityDate} · {t.presenter}: {r.presenter||"-"} · {t.totalParticipants}: <b>{r.totalParticipants}</b></div>
              {r.summary&&<div style={{fontSize:12,color:"var(--color-text-secondary)",marginTop:5}}>{r.summary}</div>}
              {r.participantNames&&<div style={{fontSize:11,color:"var(--color-text-tertiary)",marginTop:3}}>{t.participantNames}: {r.participantNames}</div>}
              <div style={{display:"flex",gap:8,marginTop:5,fontSize:11,flexWrap:"wrap"}}>
                {r.mediaStatus&&r.mediaStatus!=="Not contacted"&&<span style={{padding:"2px 7px",borderRadius:6,background:r.mediaStatus==="Published"?"#d5f5e3":r.mediaStatus==="Contacted"?"#eaf2f8":r.mediaStatus==="Declined"?"#fdedec":"#fef5e7",color:r.mediaStatus==="Published"?"#196f3d":r.mediaStatus==="Contacted"?"#2471a3":r.mediaStatus==="Declined"?"#c0392b":"#b7950b",fontWeight:600}}>{r.mediaStatus==="Published"?"📰":"📞"} {r.mediaStatus}</span>}
                {!r.mediaStatus&&r.mediaNotified&&<span style={{color:"#27ae60"}}>✅ {t.mediaNotified}</span>}
                {r.mediaChannel&&<span style={{color:"#8e44ad"}}>📺 {r.mediaChannel}</span>}
                {r.confirmed&&<span style={{color:"#27ae60"}}>✅ {t.confirmed}</span>}
              </div>
            </div>
            <div style={{display:"flex",gap:3,alignItems:"flex-start"}}>
              <button onClick={()=>printReport(r)} title={t.print} style={btnI}><Ic n="printer" s={14}/></button>
              {(isAdmin||canEdit)&&<button onClick={()=>setEditId(r.id)} style={btnI}><Ic n="edit" s={14}/></button>}
              {isAdmin&&<button onClick={()=>del(r.id)} style={btnI}><Ic n="trash" s={14} c="#e74c3c"/></button>}
            </div>
          </div>
        </div>
      ))}
      {filteredReports.length===0&&<div style={{textAlign:"center",padding:50,color:"var(--color-text-tertiary)"}}>{t.noReports}</div>}
    </div>
  );
}

function RptForm({t,allActs,onSave,onCancel,COLLEGES,initial}){
  const [d,sD]=useState({activityId:initial?.activityId||"",activityTitle:initial?.activityTitle||"",activityDate:initial?.activityDate||"",activityType:initial?.activityType||"",college:initial?.college||"",presenter:initial?.presenter||"",source:initial?.source||"",totalParticipants:initial?.totalParticipants||0,participantNames:initial?.participantNames||"",summary:initial?.summary||"",mediaNotified:initial?.mediaNotified||false,mediaLinks:initial?.mediaLinks||"",confirmed:initial?.confirmed||false,mediaStatus:initial?.mediaStatus||"Not contacted",mediaChannel:initial?.mediaChannel||"",mediaContactedDate:initial?.mediaContactedDate||"",mediaPublishedDate:initial?.mediaPublishedDate||"",mediaCoverageNotes:initial?.mediaCoverageNotes||""});
  const s=(k,v)=>sD({...d,[k]:v});
  const selAct=id=>{const f=allActs.find(a=>a.id===id);if(f)sD(prev=>({...prev,activityId:id,activityTitle:f.title,activityDate:f.date||"",activityType:f.type||"",college:f.college||"",presenter:f.presenter||"",source:f.source||"",totalParticipants:f.attendees||0,participantNames:f.participants||"",mediaNotified:f.mediaStatus==="Published"||f.mediaStatus==="Contacted"||prev.mediaNotified,mediaLinks:f.mediaLinks||prev.mediaLinks,mediaStatus:f.mediaStatus||prev.mediaStatus,mediaChannel:f.mediaChannel||prev.mediaChannel,mediaContactedDate:f.mediaContactedDate||prev.mediaContactedDate,mediaPublishedDate:f.mediaPublishedDate||prev.mediaPublishedDate,mediaCoverageNotes:f.mediaCoverageNotes||prev.mediaCoverageNotes}));};
  const [extracting,setExtracting]=useState(false);const [extractMsg,setExtractMsg]=useState("");

  const handleExtract=async(file)=>{
    if(!file)return;setExtracting(true);setExtractMsg("Extracting...");
    try{
      const dataUrl=await new Promise((res,rej)=>{const r=new FileReader();r.onload=()=>res(r.result);r.onerror=rej;r.readAsDataURL(file);});
      const [meta,b64]=dataUrl.split(",");const mtype=meta.match(/:(.*?);/)?.[1]||"image/jpeg";
      const result=await extractEventReport(b64,mtype,t.dir==="rtl"?"ar":"en");
      if(result){
        sD(prev=>({...prev,
          activityTitle:result.activityTitle||prev.activityTitle,activityDate:result.activityDate||prev.activityDate,
          activityType:result.activityType||prev.activityType,college:result.college||prev.college,
          presenter:result.presenter||prev.presenter,totalParticipants:result.totalParticipants||prev.totalParticipants,
          participantNames:result.participantNames||prev.participantNames,summary:result.summary||prev.summary,
          mediaNotified:result.mediaNotified??prev.mediaNotified,mediaLinks:result.mediaLinks||prev.mediaLinks
        }));
        // Try to match extracted title to an existing activity
        if(result.activityTitle){
          const words=result.activityTitle.toLowerCase().split(/\s+/).filter(w=>w.length>3);
          const match=allActs.find(a=>{const at=a.title.toLowerCase();return words.filter(w=>at.includes(w)).length>=Math.max(1,Math.floor(words.length/2));});
          if(match)sD(prev=>({...prev,activityId:match.id,source:match.source||""}));
        }
        setExtractMsg("Extracted successfully");
      }else setExtractMsg("Could not extract — fill manually");
    }catch(e){setExtractMsg("Error: "+e.message);}
    setExtracting(false);setTimeout(()=>setExtractMsg(""),3000);
  };

  return(
    <div style={{...crd,border:"1.5px solid #c9a84c",marginBottom:16}}>
      <h3 style={{fontSize:14,fontWeight:600,margin:"0 0 6px"}}>{initial?t.edit+" "+t.reports:t.newReport}</h3>
      <div style={{fontSize:11,color:"#27ae60",marginBottom:8,padding:"6px 10px",background:"#d5f5e3",borderRadius:6}}>ℹ️ {t.reportSyncsActivity}</div>
      {/* AI EXTRACT */}
      <div style={{display:"flex",gap:6,marginBottom:12,flexWrap:"wrap",alignItems:"center"}}>
        <label style={{...btnP,padding:"6px 14px",fontSize:11,display:"inline-flex",alignItems:"center",gap:4,cursor:extracting?"wait":"pointer",background:"#8e44ad",opacity:extracting?0.5:1}}>
          📄 {t.extractFromFile}<input type="file" accept="image/*,application/pdf" onChange={e=>{handleExtract(e.target.files[0]);e.target.value="";}} style={{display:"none"}} disabled={extracting}/>
        </label>
        <label style={{...btnP,padding:"6px 14px",fontSize:11,display:"inline-flex",alignItems:"center",gap:4,cursor:extracting?"wait":"pointer",background:"#2980b9",opacity:extracting?0.5:1}}>
          📷 Camera<input type="file" accept="image/*" capture="environment" onChange={e=>{handleExtract(e.target.files[0]);e.target.value="";}} style={{display:"none"}} disabled={extracting}/>
        </label>
        {extracting&&<span style={{fontSize:11,color:"#8e44ad"}}>⏳ Extracting...</span>}
        {extractMsg&&!extracting&&<span style={{fontSize:11,color:extractMsg.includes("success")?"#27ae60":"#e67e22"}}>{extractMsg}</span>}
      </div>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>
        <div style={{gridColumn:"1/-1"}}><label style={lbl}>{t.selectActivity}</label>
          <select value="" onChange={e=>selAct(e.target.value)} style={inp}>
            <option value="">-- {t.selectActivity} --</option>
            {(()=>{const groups={};allActs.forEach(a=>{const src=a.source||"Other";if(!groups[src])groups[src]=[];groups[src].push(a);});return Object.entries(groups).map(([src,acts])=><optgroup key={src} label={`━━ ${src} (${acts.length}) ━━`}>{acts.map(a=><option key={a.id} value={a.id}>{a.title} ({a.date||""}) {a.college?`· ${a.college.replace("College of ","").replace("Department of ","")}`:""}</option>)}</optgroup>);})()}
          </select>
        </div>
        {d.activityTitle&&<div style={{gridColumn:"1/-1",padding:"8px 12px",background:"#eaf2f8",borderRadius:6,fontSize:12}}>📋 <b>{d.activityTitle}</b> {d.activityType&&`· ${d.activityType}`} {d.source&&`· ${d.source}`} {d.college&&`· ${d.college.replace("College of ","")}`} {d.presenter&&`· ${d.presenter}`}</div>}
        <div><label style={lbl}>{t.date}</label><input type="date" value={d.activityDate} onChange={e=>s("activityDate",e.target.value)} style={inp}/></div>
        <div><label style={lbl}>{t.college}</label><select value={d.college} onChange={e=>s("college",e.target.value)} style={inp}><option value="">--</option>{(COLLEGES||[]).map(c=><option key={c}>{c}</option>)}</select></div>
        <div><label style={lbl}>{t.presenter}</label><input value={d.presenter} onChange={e=>s("presenter",e.target.value)} style={inp}/></div>
        <div><label style={lbl}>{t.totalParticipants}</label><input type="number" value={d.totalParticipants} onChange={e=>s("totalParticipants",Number(e.target.value))} style={inp}/></div>
        <div style={{gridColumn:"1/-1"}}><label style={lbl}>{t.participantNames}</label><textarea value={d.participantNames} onChange={e=>s("participantNames",e.target.value)} rows={3} style={{...inp,resize:"vertical"}}/></div>
        <div style={{gridColumn:"1/-1"}}><label style={lbl}>{t.summary}</label><textarea value={d.summary} onChange={e=>s("summary",e.target.value)} rows={2} style={{...inp,resize:"vertical"}}/></div>
        <div style={{gridColumn:"1/-1",borderTop:"1px dashed var(--color-border-tertiary)",paddingTop:10,marginTop:4}}><span style={{fontSize:12,fontWeight:600,color:"#e67e22"}}>📺 {t.mediaCoverage}</span></div>
        <div><label style={lbl}>{t.mediaStatus}</label><select value={d.mediaStatus} onChange={e=>{s("mediaStatus",e.target.value);if(e.target.value==="Published"||e.target.value==="Contacted")sD(prev=>({...prev,mediaStatus:e.target.value,mediaNotified:true}));}} style={inp}><option value="Not contacted">{t.notContacted}</option><option value="Contacted">{t.contacted}</option><option value="Coverage scheduled">{t.coverageScheduled}</option><option value="Published">{t.published}</option><option value="Declined">{t.declined}</option></select></div>
        <div><label style={lbl}>{t.mediaChannel}</label><input value={d.mediaChannel} onChange={e=>s("mediaChannel",e.target.value)} placeholder="Facebook, YouTube, Website..." style={inp}/></div>
        <div><label style={lbl}>{t.mediaContactedDate}</label><input type="date" value={d.mediaContactedDate} onChange={e=>s("mediaContactedDate",e.target.value)} style={inp}/></div>
        <div><label style={lbl}>{t.mediaPublishedDate}</label><input type="date" value={d.mediaPublishedDate} onChange={e=>s("mediaPublishedDate",e.target.value)} style={inp}/></div>
        <div style={{gridColumn:"1/-1"}}><label style={lbl}>{t.mediaLinks}</label><input value={d.mediaLinks} onChange={e=>s("mediaLinks",e.target.value)} placeholder="https://..." style={inp}/></div>
        <div style={{gridColumn:"1/-1"}}><label style={lbl}>{t.mediaCoverageNotes}</label><textarea value={d.mediaCoverageNotes} onChange={e=>s("mediaCoverageNotes",e.target.value)} rows={2} placeholder={t.mediaCoverageNotes+"..."} style={{...inp,resize:"vertical"}}/></div>
        <div><label style={{...lbl,display:"flex",alignItems:"center",gap:6}}><input type="checkbox" checked={d.confirmed} onChange={e=>s("confirmed",e.target.checked)}/>{t.confirmed}</label></div>
      </div>
      <div style={{display:"flex",gap:8,marginTop:12,justifyContent:"flex-end"}}><button onClick={onCancel} style={btnC}>{t.cancel}</button><button onClick={()=>d.activityTitle&&onSave(d)} style={{...btnP,opacity:d.activityTitle?1:0.5}}>{t.save}</button></div>
    </div>
  );
}

// ─── CORRESPONDENCE ───────────────────────────────────────────────────
function CorrPage({t,corr,saveCorr,isAdmin,canEdit,COLLEGES}){
  const [sf,sSf]=useState(false);const [flt,sFlt]=useState("All");const [editId,setEditId]=useState(null);
  const [fCollege,setFCollege]=useState("All");const [fSearch,setFSearch]=useState("");
  const [bulkExtracting,setBE]=useState(false);const [bulkResults,setBR]=useState(null);const [bulkMsg,setBM]=useState("");
  const add=d=>{saveCorr([...corr,{id:gid(),...d}]);sSf(false);};
  const upd=d=>{saveCorr(corr.map(c=>c.id===editId?{...c,...d}:c));setEditId(null);};
  const addBulk=items=>{const newItems=items.map(i=>({id:gid(),type:i.type==="Outgoing"?"Outgoing":"Incoming",refNumber:i.refNumber||"",date:i.date||"",from:i.from||"",to:i.to||"",subject:i.subject||"",notes:i.notes||"",status:i.status||"Received",college:i.college||"",customFields:i.customFields||[],fileData:"",fileName:""}));saveCorr([...corr,...newItems]);setBR(null);setBM("");};
  const del=id=>saveCorr(corr.filter(c=>c.id!==id));
  const clearFile=id=>saveCorr(corr.map(c=>c.id===id?{...c,fileData:"",fileName:""}:c));
  const filtered=corr.filter(c=>{
    if(flt!=="All"&&c.type!==flt)return false;
    if(fCollege!=="All"&&c.college!==fCollege)return false;
    if(fSearch){
      const q=fSearch.toLowerCase();
      const searchable=[c.subject,c.refNumber,c.from,c.to,c.notes,c.date,c.status,c.college,...(c.customFields||[]).map(f=>f.name+" "+f.value)].filter(Boolean).join(" ").toLowerCase();
      if(!searchable.includes(q))return false;
    }
    return true;
  });
  const corrColleges=[...new Set(corr.map(c=>c.college).filter(Boolean))];
  const handleBulkFile=async file=>{
    if(!file)return;setBE(true);setBM("");setBR(null);
    if(file.size>4*1024*1024){setBM("⚠️ Max 4MB");setBE(false);return;}
    const mtype=file.type||"";if(!mtype.startsWith("image/")&&mtype!=="application/pdf"&&!mtype.startsWith("text/")){setBM("⚠️ Use PDF, image, or camera");setBE(false);return;}
    try{const b64=await new Promise((r,j)=>{const rd=new FileReader();rd.onload=()=>r(rd.result.split(",")[1]);rd.onerror=()=>j();rd.readAsDataURL(file);});
      const items=await extractBulkCorrespondence(b64,mtype,t.dir==="rtl"?"ar":"en");
      if(items?.length>0){setBR(items);setBM(`✅ ${items.length} items found`);}else{setBM("❌ "+t.extractionError);}
    }catch{setBM("❌ "+t.extractionError);}setBE(false);
  };

  return(
    <div>
      <div style={{display:"flex",justifyContent:"space-between",marginBottom:16}}>
        <div><h2 style={{fontSize:18,fontWeight:700,margin:0}}>{t.correspondence}</h2><p style={{fontSize:12,color:"var(--color-text-secondary)",margin:"3px 0 0"}}>{t.corrSubtitle}</p></div>
        {(isAdmin||canEdit)&&<button onClick={()=>sSf(true)} style={{...btnP,display:"flex",alignItems:"center",gap:5}}><Ic n="plus" s={15} c="#fff"/>{t.add}</button>}
      </div>
      <div style={{display:"flex",gap:8,marginBottom:8,flexWrap:"wrap"}}>{[t.all,t.incomingCorr,t.outgoingCorr].map((f,i)=>{const v=["All","Incoming","Outgoing"][i];return <button key={v} onClick={()=>sFlt(v)} style={{padding:"6px 14px",borderRadius:20,border:"1px solid var(--color-border-tertiary)",background:flt===v?"#1a5276":"transparent",color:flt===v?"#fff":"var(--color-text-secondary)",fontSize:12,cursor:"pointer",fontWeight:flt===v?600:400}}>{f}</button>;})}</div>
      <div style={{display:"flex",gap:8,marginBottom:14,flexWrap:"wrap",alignItems:"center"}}>
        <select value={fCollege} onChange={e=>setFCollege(e.target.value)} style={{...inp,width:"auto",minWidth:160,padding:"5px 10px",fontSize:11}}>
          <option value="All">🏛️ {t.all} {t.college}</option>
          {(COLLEGES||corrColleges).map(c=><option key={c}>{c}</option>)}
        </select>
        <input value={fSearch} onChange={e=>setFSearch(e.target.value)} placeholder={"🔍 "+t.search+"..."} style={{...inp,width:"auto",minWidth:180,padding:"5px 10px",fontSize:11}}/>
        <span style={{fontSize:11,color:"var(--color-text-tertiary)"}}>{filtered.length}/{corr.length}</span>
      </div>
      {sf&&<CrForm t={t} onSave={add} onCancel={()=>sSf(false)} COLLEGES={COLLEGES}/>}
      {editId&&<CrForm t={t} initial={corr.find(c=>c.id===editId)} onSave={upd} onCancel={()=>setEditId(null)} COLLEGES={COLLEGES}/>}
      {/* BULK IMPORT */}
      {(isAdmin||canEdit)&&<div style={{background:"var(--color-background-secondary)",borderRadius:10,padding:14,marginBottom:14,border:"1px dashed var(--color-border-secondary)"}}>
        <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:6}}><span style={{fontSize:16}}>📋</span><div><div style={{fontSize:12,fontWeight:600}}>{t.bulkCorrImport}</div><div style={{fontSize:10,color:"var(--color-text-tertiary)"}}>{t.bulkCorrHint}</div></div></div>
        <div style={{display:"flex",gap:8,alignItems:"center",flexWrap:"wrap"}}>
          <label style={{...btnP,padding:"5px 12px",fontSize:11,cursor:"pointer",opacity:bulkExtracting?0.5:1}}><Ic n="plan" s={13} c="#fff"/>&nbsp;{t.fileUpload}<input type="file" accept="image/*,.pdf,.txt" onChange={e=>handleBulkFile(e.target.files[0])} style={{display:"none"}} disabled={bulkExtracting}/></label>
          <label style={{...btnP,padding:"5px 12px",fontSize:11,cursor:"pointer",background:"#27ae60",opacity:bulkExtracting?0.5:1}}>📷 {t.camera}<input type="file" accept="image/*" capture="environment" onChange={e=>handleBulkFile(e.target.files[0])} style={{display:"none"}} disabled={bulkExtracting}/></label>
          {bulkExtracting&&<span style={{fontSize:11,color:"#e67e22"}}>⏳ {t.extracting}</span>}
        </div>
        {bulkMsg&&<div style={{marginTop:6,fontSize:11,color:bulkResults?"#27ae60":"#e74c3c"}}>{bulkMsg}</div>}
        {bulkResults?.length>0&&<div style={{marginTop:10}}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:6}}><span style={{fontSize:11,fontWeight:600}}>{bulkResults.length} items:</span><button onClick={()=>addBulk(bulkResults)} style={{...btnP,padding:"4px 12px",fontSize:11}}>{t.importAll} ({bulkResults.length})</button></div>
          {bulkResults.map((it,i)=><div key={i} style={{background:"#ffffff",borderRadius:6,padding:"6px 10px",marginBottom:4,border:"0.5px solid var(--color-border-tertiary)",fontSize:11,display:"flex",justifyContent:"space-between"}}><div><b>{it.subject||`Item ${i+1}`}</b> · {it.type} · {it.date} · {it.from}→{it.to}</div><button onClick={()=>setBR(bulkResults.filter((_,idx)=>idx!==i))} style={{background:"none",border:"none",cursor:"pointer",color:"#e74c3c",fontSize:11}}>✕</button></div>)}
        </div>}
      </div>}
      {filtered.map(c=>(
        <div key={c.id} style={{...crd,marginBottom:8,padding:"12px 16px",borderLeft:`3px solid ${c.type==="Incoming"?"#3498db":"#e67e22"}`}}>
          <div style={{display:"flex",justifyContent:"space-between"}}>
            <div style={{flex:1,minWidth:0}}>
              <div style={{display:"flex",gap:6,alignItems:"center",marginBottom:3,flexWrap:"wrap"}}>
                <span style={{fontSize:13}}>{c.type==="Incoming"?"📥":"📤"}</span>
                <span style={{fontSize:13,fontWeight:600}}>{c.subject}</span>
                <span style={{fontSize:9,padding:"2px 7px",borderRadius:10,background:c.type==="Incoming"?"#eaf2f8":"#fef5e7",color:c.type==="Incoming"?"#2471a3":"#ca6f1e",fontWeight:600}}>{c.type==="Incoming"?t.incomingCorr:t.outgoingCorr}</span>
                <span style={{fontSize:9,padding:"2px 7px",borderRadius:10,background:"var(--color-background-secondary)",color:"var(--color-text-secondary)"}}>{c.status}</span>
                {c.college&&<span style={{fontSize:9,padding:"2px 7px",borderRadius:10,background:"#e8d5f5",color:"#6c3483"}}>{c.college.replace("College of ","").replace("Department of ","")}</span>}
              </div>
              <div style={{fontSize:11,color:"var(--color-text-secondary)"}}>{t.refNumber}: <b>{c.refNumber}</b> · {t.date}: {c.date} · {t.from}: {c.from} → {t.to}: {c.to}</div>
              {c.notes&&<div style={{fontSize:11,color:"var(--color-text-tertiary)",marginTop:3,fontStyle:"italic"}}>{c.notes}</div>}
              {c.customFields&&c.customFields.length>0&&<div style={{marginTop:4,display:"flex",gap:6,flexWrap:"wrap"}}>{c.customFields.filter(f=>f.name).map((f,i)=><span key={i} style={{fontSize:10,padding:"2px 8px",borderRadius:8,background:"#f0e6f6",color:"#6c3483",border:"0.5px solid #d7bde2"}}>{f.name}: {f.value}</span>)}</div>}
              {c.fileName&&<div style={{marginTop:6,display:"flex",gap:6,alignItems:"center"}}>
                <button onClick={()=>downloadFile(c.fileData,c.fileName)} style={{display:"inline-flex",alignItems:"center",gap:5,padding:"4px 10px",borderRadius:6,border:"1px solid var(--color-border-tertiary)",background:"var(--color-background-secondary)",cursor:"pointer",fontSize:11,color:"#2980b9"}}>
                  <Ic n="download" s={13} c="#2980b9"/>{c.fileName}
                </button>
                {(isAdmin||canEdit)&&<button onClick={()=>clearFile(c.id)} style={{display:"inline-flex",alignItems:"center",gap:3,padding:"4px 8px",borderRadius:6,border:"1px solid #e74c3c",background:"transparent",cursor:"pointer",fontSize:10,color:"#e74c3c"}}>✕ {t.removeField}</button>}
              </div>}
            </div>
            {(isAdmin||canEdit)&&<div style={{display:"flex",gap:3,flexShrink:0,alignItems:"flex-start"}}>
              <button onClick={()=>setEditId(c.id)} style={btnI}><Ic n="edit" s={14}/></button>
              <button onClick={()=>del(c.id)} style={btnI}><Ic n="trash" s={14} c="#e74c3c"/></button>
            </div>}
          </div>
        </div>
      ))}
      {filtered.length===0&&<div style={{textAlign:"center",padding:50,color:"var(--color-text-tertiary)"}}>{t.noCorr}</div>}
    </div>
  );
}

function CrForm({t,onSave,onCancel,initial,COLLEGES}){
  const [d,sD]=useState({type:initial?.type||"Incoming",refNumber:initial?.refNumber||"",date:initial?.date||new Date().toISOString().split("T")[0],from:initial?.from||"",to:initial?.to||"",subject:initial?.subject||"",notes:initial?.notes||"",status:initial?.status||"Received",college:initial?.college||""});
  const [customFields,setCF]=useState(initial?.customFields||[]);
  const [fileData,setFileData]=useState(initial?.fileData||"");
  const [fileName,setFileName]=useState(initial?.fileName||"");
  const [extracting,setExtracting]=useState(false);
  const [extractMsg,setExtractMsg]=useState("");
  const s=(k,v)=>sD({...d,[k]:v});
  const addCF=()=>setCF([...customFields,{id:gid(),name:"",value:""}]);
  const updCF=(id,field,val)=>setCF(customFields.map(f=>f.id===id?{...f,[field]:val}:f));
  const delCF=id=>setCF(customFields.filter(f=>f.id!==id));

  const handleFile=async(file)=>{
    if(!file)return;
    const {dataUrl:compressed, originalDataUrl}=await processFileForStorage(file);
    setFileData(compressed);
    setFileName(file.name);

    if(file.size>4*1024*1024){setExtractMsg("✅ File attached (compressed).");return;}
    const mtype=file.type||"application/octet-stream";
    const isSupported=mtype.startsWith("image/")||mtype==="application/pdf"||mtype.startsWith("text/");
    if(!isSupported){setExtractMsg("✅ File attached. (.docx not supported for extraction — use PDF or image)");return;}

    setExtracting(true);setExtractMsg("");
    try{
      const b64=originalDataUrl.split(",")[1];
      const result=await extractCorrespondence(b64,mtype,t.dir==="rtl"?"ar":"en");
      if(result){
        const newD={...d};
        ["type","refNumber","date","from","to","subject","notes","status"].forEach(k=>{
          if(result[k]&&typeof result[k]==="string"&&result[k].trim())newD[k]=result[k].trim();
        });
        if(result.type==="Incoming"||result.type==="Outgoing")newD.type=result.type;
        sD(newD);
        if(result.customFields&&Array.isArray(result.customFields)&&result.customFields.length>0){
          setCF(prev=>[...prev,...result.customFields.map(cf=>({id:gid(),name:cf.name||"",value:cf.value||""}))]);
        }
        setExtractMsg(t.extractionDone);
      }else{setExtractMsg("✅ File attached. Extraction unavailable — fill manually.");}
    }catch(e){console.error(e);setExtractMsg("✅ File attached. Extraction error — fill manually.");}
    setExtracting(false);
  };

  const doSave=()=>{
    if(!d.subject)return;
    onSave({...d, customFields:customFields.filter(f=>f.name.trim()), fileData, fileName});
  };

  return(
    <div style={{...crd,border:"1.5px solid #c9a84c",marginBottom:16}}>
      <h3 style={{fontSize:14,fontWeight:600,margin:"0 0 8px"}}>{t.newCorr}</h3>

      {/* FILE UPLOAD */}
      <div style={{background:"var(--color-background-secondary)",borderRadius:10,padding:14,marginBottom:14,border:"1px dashed var(--color-border-secondary)"}}>
        <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:8}}>
          <span style={{fontSize:18}}>📄</span>
          <div><div style={{fontSize:12,fontWeight:600}}>{t.extractFromFile}</div><div style={{fontSize:10,color:"var(--color-text-tertiary)"}}>{t.corrExtractHint}</div></div>
        </div>
        <div style={{display:"flex",gap:10,alignItems:"center",flexWrap:"wrap"}}>
          <label style={{...btnP,padding:"6px 14px",fontSize:11,display:"inline-flex",alignItems:"center",gap:4,cursor:"pointer",opacity:extracting?0.5:1}}>
            <Ic n="plan" s={13} c="#fff"/>{t.fileUpload}
            <input type="file" accept="image/*,.pdf,.txt" onChange={e=>handleFile(e.target.files[0])} style={{display:"none"}} disabled={extracting}/>
          </label>
          <label style={{...btnP,padding:"6px 14px",fontSize:11,display:"inline-flex",alignItems:"center",gap:4,cursor:"pointer",background:"#27ae60",opacity:extracting?0.5:1}}>
            📷 {t.camera}
            <input type="file" accept="image/*" capture="environment" onChange={e=>handleFile(e.target.files[0])} style={{display:"none"}} disabled={extracting}/>
          </label>
          {extracting&&<span style={{fontSize:11,color:"#e67e22",fontWeight:600}}>⏳ {t.extracting}</span>}
          {fileName&&<span style={{fontSize:11,color:"#27ae60"}}>📎 {fileName}</span>}
        </div>
        {extractMsg&&<div style={{marginTop:6,fontSize:11,color:extractMsg.includes("✅")||extractMsg===t.extractionDone?"#27ae60":"#e74c3c"}}>{extractMsg}</div>}
      </div>

      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:10}}>
        <div><label style={lbl}>{t.type}</label><select value={d.type} onChange={e=>s("type",e.target.value)} style={inp}><option value="Incoming">{t.incoming}</option><option value="Outgoing">{t.outgoing}</option></select></div>
        <div><label style={lbl}>{t.refNumber}</label><input value={d.refNumber} onChange={e=>s("refNumber",e.target.value)} style={inp}/></div>
        <div><label style={lbl}>{t.date}</label><input type="date" value={d.date} onChange={e=>s("date",e.target.value)} style={inp}/></div>
        <div><label style={lbl}>{t.from}</label><input value={d.from} onChange={e=>s("from",e.target.value)} style={inp}/></div>
        <div><label style={lbl}>{t.to}</label><input value={d.to} onChange={e=>s("to",e.target.value)} style={inp}/></div>
        <div><label style={lbl}>{t.status}</label><select value={d.status} onChange={e=>s("status",e.target.value)} style={inp}><option>{t.received}</option><option>{t.processed}</option><option>{t.sent}</option><option>{t.pending}</option><option>{t.archived}</option></select></div>
        <div><label style={lbl}>{t.college}</label><select value={d.college} onChange={e=>s("college",e.target.value)} style={inp}><option value="">— {t.all} —</option>{(COLLEGES||[]).map(c=><option key={c}>{c}</option>)}</select></div>
        <div style={{gridColumn:"1/-1"}}><label style={lbl}>{t.subject}</label><input value={d.subject} onChange={e=>s("subject",e.target.value)} style={inp}/></div>
        <div style={{gridColumn:"1/-1"}}><label style={lbl}>{t.notes}</label><textarea value={d.notes} onChange={e=>s("notes",e.target.value)} rows={2} style={{...inp,resize:"vertical"}}/></div>

        {/* CUSTOM FIELDS */}
        {customFields.length>0&&<div style={{gridColumn:"1/-1",borderTop:"1px dashed var(--color-border-tertiary)",paddingTop:8,marginTop:4}}><span style={{fontSize:12,fontWeight:600,color:"#8e44ad"}}>📋 {t.extractedFields}</span></div>}
        {customFields.map(cf=>(
          <div key={cf.id} style={{gridColumn:"1/-1",display:"flex",gap:8,alignItems:"flex-start"}}>
            <div style={{flex:1}}><label style={lbl}>{t.fieldName}</label><input value={cf.name} onChange={e=>updCF(cf.id,"name",e.target.value)} style={inp}/></div>
            <div style={{flex:2}}><label style={lbl}>{t.fieldValue}</label><input value={cf.value} onChange={e=>updCF(cf.id,"value",e.target.value)} style={inp}/></div>
            <button onClick={()=>delCF(cf.id)} style={{...btnI,marginTop:18}} title={t.removeField}><Ic n="trash" s={14} c="#e74c3c"/></button>
          </div>
        ))}
        <div style={{gridColumn:"1/-1"}}><button onClick={addCF} style={{...btnC,fontSize:11,padding:"5px 12px",display:"inline-flex",alignItems:"center",gap:4}}><Ic n="plus" s={13}/>{t.addCustomField}</button></div>
      </div>
      <div style={{display:"flex",gap:8,marginTop:12,justifyContent:"flex-end"}}><button onClick={onCancel} style={btnC}>{t.cancel}</button><button onClick={doSave} style={{...btnP,opacity:d.subject?1:0.5}}>{t.save}</button></div>
    </div>
  );
}

// ─── SETTINGS ─────────────────────────────────────────────────────────
function SettingsPage({t,config,saveConfig,users,saveUsers,currentUser,audit,COLLEGES,plans,savePlans,inits,saveInits,trainers,saveTrainers,corr,saveCorr,reports,saveReports}){
  const [pn,sPn]=useState(config.platformName||"");
  const [ps,sPs]=useState(config.platformSub||"");
  const [showAddUser,sShowAddUser]=useState(false);
  const [nu,sNu]=useState({username:"",password:"",name:"",role:"viewer",college:""});
  const [cpCurrent,setCpCurrent]=useState("");
  const [cpNew,setCpNew]=useState("");
  const [cpConfirm,setCpConfirm]=useState("");
  const [cpMsg,setCpMsg]=useState("");
  const [apiKey,setApiKey]=useState(()=>{try{return localStorage.getItem("ced_api_key")||"";}catch{return "";}});

  const saveNames=()=>saveConfig({...config,platformName:pn,platformSub:ps});
  const addUser=async()=>{
    if(!nu.username||!nu.password||!nu.name)return;
    const hash=await hashPassword(nu.password);
    saveUsers([...users,{id:gid(),username:nu.username,name:nu.name,role:nu.role,college:nu.college||null,passwordHash:hash}]);
    if(audit)audit(t.userCreated,nu.username);
    sNu({username:"",password:"",name:"",role:"viewer",college:""});sShowAddUser(false);
  };
  const delUser=id=>{const u=users.find(x=>x.id===id);saveUsers(users.filter(x=>x.id!==id));if(audit&&u)audit(t.userDeleted,u.username);};
  const updUser=(id,field,val)=>saveUsers(users.map(u=>u.id===id?{...u,[field]:val}:u));

  const changePassword=async()=>{
    setCpMsg("");
    if(cpNew!==cpConfirm){setCpMsg(t.passwordMismatch);return;}
    if(cpNew.length<4){setCpMsg("Password too short");return;}
    const me=users.find(u=>u.username===currentUser?.username);
    if(!me){setCpMsg("User not found");return;}
    // Verify current password
    let valid=false;
    if(me.passwordHash)valid=await verifyPassword(cpCurrent,me.passwordHash);
    else if(me.password)valid=(me.password===cpCurrent);
    if(!valid){setCpMsg(t.wrongPassword);return;}
    const hash=await hashPassword(cpNew);
    saveUsers(users.map(u=>u.username===currentUser.username?{...u,passwordHash:hash,password:undefined}:u));
    setCpMsg("✅ "+t.passwordChanged);setCpCurrent("");setCpNew("");setCpConfirm("");
    if(audit)audit(t.changePassword,"");
  };

  return(
    <div>
      <h2 style={{fontSize:18,fontWeight:700,margin:"0 0 3px"}}>{t.settings}</h2>
      <p style={{fontSize:12,color:"var(--color-text-secondary)",margin:"0 0 16px"}}>{t.settingsSubtitle}</p>

      {/* Platform name */}
      <div style={{...crd,marginBottom:16}}>
        <h3 style={{fontSize:14,fontWeight:600,margin:"0 0 12px"}}>{t.platformName}</h3>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>
          <div><label style={lbl}>{t.platformName}</label><input value={pn} onChange={e=>sPn(e.target.value)} placeholder={t.appDefaultName} style={inp}/></div>
          <div><label style={lbl}>{t.platformSub}</label><input value={ps} onChange={e=>sPs(e.target.value)} placeholder={t.appDefaultSub} style={inp}/></div>
        </div>
        <button onClick={saveNames} style={{...btnP,marginTop:12}}>{t.save}</button>
      </div>

      {/* Language */}
      <div style={{...crd,marginBottom:16}}>
        <h3 style={{fontSize:14,fontWeight:600,margin:"0 0 12px"}}>{t.language}</h3>
        <div style={{display:"flex",gap:10}}>
          <button onClick={()=>saveConfig({...config,lang:"en"})} style={{padding:"8px 20px",borderRadius:8,border:config.lang==="en"?"2px solid #1a5276":"1px solid var(--color-border-tertiary)",background:config.lang==="en"?"#eaf2f8":"transparent",cursor:"pointer",fontWeight:600}}>English</button>
          <button onClick={()=>saveConfig({...config,lang:"ar"})} style={{padding:"8px 20px",borderRadius:8,border:config.lang==="ar"?"2px solid #1a5276":"1px solid var(--color-border-tertiary)",background:config.lang==="ar"?"#eaf2f8":"transparent",cursor:"pointer",fontWeight:600,fontFamily:"'Noto Sans Arabic',sans-serif"}}>العربية</button>
        </div>
      </div>

      {/* AI API Key */}
      <div style={{...crd,marginBottom:16}}>
        <h3 style={{fontSize:14,fontWeight:600,margin:"0 0 8px"}}>{t.dir==="rtl"?"مفتاح Claude API":"Claude API Key"}</h3>
        <p style={{fontSize:11,color:"var(--color-text-tertiary)",margin:"0 0 10px"}}>{t.dir==="rtl"?"مطلوب لميزات الاستخراج بالذكاء الاصطناعي على الموقع المنشور":"Required for AI extraction features on deployed site"}</p>
        <div style={{display:"flex",gap:8,alignItems:"center"}}>
          <input type="password" placeholder="sk-ant-..." value={apiKey} onChange={e=>{const v=e.target.value;setApiKey(v);const trimmed=v.trim();if(trimmed)localStorage.setItem("ced_api_key",trimmed);else localStorage.removeItem("ced_api_key");}} style={{...inp,flex:1,fontFamily:"monospace",fontSize:12}}/>
          {apiKey.trim()&&<span style={{color:"#27ae60",fontSize:20,fontWeight:700}}>✅</span>}
        </div>
      </div>

      {/* Logo upload */}
      <div style={{...crd,marginBottom:16}}>
        <h3 style={{fontSize:14,fontWeight:600,margin:"0 0 8px"}}>{t.platformLogo}</h3>
        <p style={{fontSize:11,color:"var(--color-text-tertiary)",margin:"0 0 12px"}}>{t.logoHelp}</p>
        <div style={{display:"flex",alignItems:"center",gap:16}}>
          <img src={config.customLogo||LOGO} alt="Logo" style={{width:60,height:60,borderRadius:"50%",objectFit:"contain",background:"#0d3249",border:"2px solid #c9a84c"}}/>
          <div>
            <input type="file" accept="image/png,image/jpeg" onChange={e=>{const f=e.target.files[0];if(!f)return;const r=new FileReader();r.onload=ev=>{saveConfig({...config,customLogo:ev.target.result});};r.readAsDataURL(f);}} style={{fontSize:12,marginBottom:8}}/>
            {config.customLogo&&<button onClick={()=>saveConfig({...config,customLogo:""})} style={{...btnC,padding:"4px 12px",fontSize:11}}>{t.resetLogo}</button>}
          </div>
        </div>
      </div>

      {/* Colleges & Departments management */}
      <CollegesManager t={t} config={config} saveConfig={saveConfig} COLLEGES={COLLEGES} plans={plans} savePlans={savePlans} inits={inits} saveInits={saveInits} trainers={trainers} saveTrainers={saveTrainers} corr={corr} saveCorr={saveCorr} reports={reports} saveReports={saveReports}/>

      {/* User management */}
      <div style={{...crd}}>
        <div style={{display:"flex",justifyContent:"space-between",marginBottom:12}}>
          <h3 style={{fontSize:14,fontWeight:600,margin:0}}>{t.manageUsers}</h3>
          <button onClick={()=>sShowAddUser(!showAddUser)} style={{...btnP,padding:"4px 12px",fontSize:11}}>{t.addUser}</button>
        </div>
        {showAddUser&&<div style={{background:"var(--color-background-secondary)",borderRadius:8,padding:14,marginBottom:12}}>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:10}}>
            <div><label style={lbl}>{t.username}</label><input value={nu.username} onChange={e=>sNu({...nu,username:e.target.value})} style={inp}/></div>
            <div><label style={lbl}>{t.password}</label><input value={nu.password} onChange={e=>sNu({...nu,password:e.target.value})} style={inp}/></div>
            <div><label style={lbl}>{t.name}</label><input value={nu.name} onChange={e=>sNu({...nu,name:e.target.value})} style={inp}/></div>
            <div><label style={lbl}>{t.role}</label><select value={nu.role} onChange={e=>sNu({...nu,role:e.target.value})} style={inp}><option value="admin">{t.admin}</option><option value="viewer">{t.viewer}</option><option value="college">{t.collegeRep}</option></select></div>
            {nu.role==="college"&&<div><label style={lbl}>{t.college}</label><select value={nu.college} onChange={e=>sNu({...nu,college:e.target.value})} style={inp}><option value="">--</option>{COLLEGES.map(c=><option key={c}>{c}</option>)}</select></div>}
          </div>
          <button onClick={addUser} style={{...btnP,marginTop:10}}>{t.save}</button>
        </div>}
        <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
          <thead><tr style={{borderBottom:"1px solid var(--color-border-tertiary)"}}><th style={{textAlign:"start",padding:"8px",fontWeight:600,color:"var(--color-text-secondary)"}}>{t.username}</th><th style={{textAlign:"start",padding:"8px",fontWeight:600,color:"var(--color-text-secondary)"}}>{t.name}</th><th style={{textAlign:"start",padding:"8px",fontWeight:600,color:"var(--color-text-secondary)"}}>{t.role}</th><th style={{textAlign:"start",padding:"8px",fontWeight:600,color:"var(--color-text-secondary)"}}>{t.college}</th><th style={{padding:"8px"}}></th></tr></thead>
          <tbody>{users.map(u=>(
            <tr key={u.id} style={{borderBottom:"0.5px solid var(--color-border-tertiary)"}}>
              <td style={{padding:"8px"}}><code style={{fontSize:11,background:"var(--color-background-secondary)",padding:"2px 6px",borderRadius:4}}>{u.username}</code></td>
              <td style={{padding:"8px"}}><input value={u.name} onChange={e=>updUser(u.id,"name",e.target.value)} style={{...inp,padding:"4px 8px",fontSize:12}}/></td>
              <td style={{padding:"8px"}}><select value={u.role} onChange={e=>updUser(u.id,"role",e.target.value)} style={{...inp,padding:"4px 8px",fontSize:11}}><option value="admin">{t.admin}</option><option value="viewer">{t.viewer}</option><option value="college">{t.collegeRep}</option></select></td>
              <td style={{padding:"8px",color:u.college?"inherit":"var(--color-text-tertiary)",fontSize:11}}>{u.college||t.all}</td>
              <td style={{padding:"8px"}}><button onClick={()=>delUser(u.id)} style={btnI}><Ic n="trash" s={13} c="#e74c3c"/></button></td>
            </tr>
          ))}</tbody>
        </table>
      </div>

      {/* Change Password */}
      <div style={{...crd,marginTop:16}}>
        <h3 style={{fontSize:14,fontWeight:600,margin:"0 0 12px"}}>🔐 {t.changePassword}</h3>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:10}}>
          <div><label style={lbl}>{t.currentPassword}</label><input type="password" value={cpCurrent} onChange={e=>setCpCurrent(e.target.value)} style={inp}/></div>
          <div><label style={lbl}>{t.newPassword}</label><input type="password" value={cpNew} onChange={e=>setCpNew(e.target.value)} style={inp}/></div>
          <div><label style={lbl}>{t.confirmPassword}</label><input type="password" value={cpConfirm} onChange={e=>setCpConfirm(e.target.value)} style={inp}/></div>
        </div>
        {cpMsg&&<div style={{marginTop:8,fontSize:12,color:cpMsg.includes("✅")?"#27ae60":"#e74c3c",fontWeight:500}}>{cpMsg}</div>}
        <button onClick={changePassword} style={{...btnP,marginTop:10}}>{t.changePassword}</button>
      </div>

      {/* Recovery Key Setup */}
      <RecoveryKeySetup t={t} audit={audit}/>
    </div>
  );
}

function CollegesManager({t,config,saveConfig,COLLEGES,plans,savePlans,inits,saveInits,trainers,saveTrainers,corr,saveCorr,reports,saveReports}){
  const [newName,setNewName]=useState("");
  const [msg,setMsg]=useState("");
  const [editIdx,setEditIdx]=useState(-1);
  const [editVal,setEditVal]=useState("");

  const addCollege=()=>{
    if(!newName.trim())return;
    if(COLLEGES.includes(newName.trim())){setMsg("⚠️ Already exists");return;}
    const updated=[...COLLEGES,newName.trim()];
    saveConfig({...config,colleges:updated});
    setNewName("");setMsg("✅ "+t.collegesUpdated);
    setTimeout(()=>setMsg(""),2000);
  };

  const removeCollege=(name)=>{
    const updated=COLLEGES.filter(c=>c!==name);
    saveConfig({...config,colleges:updated});
    setMsg("✅ "+t.collegesUpdated);
    setTimeout(()=>setMsg(""),2000);
  };

  const startEdit=(i)=>{setEditIdx(i);setEditVal(COLLEGES[i]);};

  const saveEdit=()=>{
    if(editIdx<0||!editVal.trim())return;
    const oldName=COLLEGES[editIdx];
    const newN=editVal.trim();
    if(oldName===newN){setEditIdx(-1);return;}
    if(COLLEGES.includes(newN)){setMsg("⚠️ Already exists");return;}
    // Update colleges list
    const updated=COLLEGES.map((c,i)=>i===editIdx?newN:c);
    saveConfig({...config,colleges:updated});
    // Propagate rename across all data
    if(plans&&savePlans)savePlans(plans.map(p=>p.college===oldName?{...p,college:newN}:p));
    if(inits&&saveInits)saveInits(inits.map(i=>i.college===oldName?{...i,college:newN}:i));
    if(trainers&&saveTrainers)saveTrainers(trainers.map(x=>x.college===oldName?{...x,college:newN}:x));
    if(corr&&saveCorr)saveCorr(corr.map(c=>c.college===oldName?{...c,college:newN}:c));
    if(reports&&saveReports)saveReports(reports.map(r=>r.college===oldName?{...r,college:newN}:r));
    setEditIdx(-1);setMsg("✅ "+t.collegesUpdated+" ("+oldName+" → "+newN+")");
    setTimeout(()=>setMsg(""),3000);
  };

  const resetToDefault=()=>{
    saveConfig({...config,colleges:undefined});
    setMsg("✅ "+t.collegesUpdated);
    setTimeout(()=>setMsg(""),2000);
  };

  return(
    <div style={{...crd,marginBottom:16}}>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:12}}>
        <h3 style={{fontSize:14,fontWeight:600,margin:0}}>🏛️ {t.manageColleges}</h3>
        <span style={{fontSize:11,color:"var(--color-text-tertiary)"}}>{COLLEGES.length} {t.collegeCount}</span>
      </div>

      {/* Add new */}
      <div style={{display:"flex",gap:8,marginBottom:12}}>
        <input value={newName} onChange={e=>setNewName(e.target.value)} onKeyDown={e=>e.key==="Enter"&&addCollege()} placeholder={t.collegeName} style={{...inp,flex:1}}/>
        <button onClick={addCollege} style={{...btnP,padding:"8px 16px",fontSize:12,whiteSpace:"nowrap"}}>{t.addCollege}</button>
      </div>
      {msg&&<div style={{fontSize:11,marginBottom:10,color:msg.includes("✅")?"#27ae60":"#e67e22"}}>{msg}</div>}

      {/* List */}
      <div style={{maxHeight:250,overflowY:"auto",border:"1px solid var(--color-border-tertiary)",borderRadius:8}}>
        {COLLEGES.map((c,i)=>(
          <div key={i} style={{display:"flex",alignItems:"center",gap:8,padding:"6px 12px",borderBottom:i<COLLEGES.length-1?"0.5px solid var(--color-border-tertiary)":"none",fontSize:12}}>
            {editIdx===i?(
              <div style={{flex:1,display:"flex",gap:6,alignItems:"center"}}>
                <input value={editVal} onChange={e=>setEditVal(e.target.value)} onKeyDown={e=>{if(e.key==="Enter")saveEdit();if(e.key==="Escape")setEditIdx(-1);}} autoFocus style={{...inp,flex:1,padding:"4px 8px",fontSize:12}}/>
                <button onClick={saveEdit} style={{...btnP,padding:"3px 10px",fontSize:10}}>✓</button>
                <button onClick={()=>setEditIdx(-1)} style={{...btnC,padding:"3px 10px",fontSize:10}}>✕</button>
              </div>
            ):(
              <>
                <span style={{flex:1,cursor:"pointer"}} onClick={()=>startEdit(i)} title={t.edit}>{c}</span>
                <button onClick={()=>startEdit(i)} style={{...btnI,padding:2}} title={t.edit}><Ic n="edit" s={12}/></button>
                <button onClick={()=>removeCollege(c)} style={{background:"none",border:"none",cursor:"pointer",color:"#e74c3c",fontSize:11,padding:"2px 4px"}} title={t.removeCollege}>✕</button>
              </>
            )}
          </div>
        ))}
      </div>

      <button onClick={resetToDefault} style={{...btnC,fontSize:11,padding:"5px 12px",marginTop:10}}>{t.resetColleges}</button>
    </div>
  );
}

function RecoveryKeySetup({t,audit}){
  const [hasKey,setHasKey]=useState(false);
  const [showKey,setShowKey]=useState("");
  const [generating,setGenerating]=useState(false);

  useEffect(()=>{(async()=>{const r=await sG(K.recovery,null);setHasKey(!!r?.hash);})();},[]);

  const generate=async()=>{
    setGenerating(true);
    const key=generateRecoveryKey();
    await saveRecoveryKeyHash(key);
    setShowKey(key);
    setHasKey(true);
    if(audit)audit("Recovery key generated","");
    setGenerating(false);
  };

  return(
    <div style={{...crd,marginTop:16}}>
      <h3 style={{fontSize:14,fontWeight:600,margin:"0 0 8px"}}>🔑 {t.setupRecovery}</h3>
      <p style={{fontSize:11,color:"var(--color-text-tertiary)",margin:"0 0 12px",lineHeight:1.7}}>{t.recoveryWarning}</p>

      {hasKey&&!showKey&&<div style={{display:"flex",alignItems:"center",gap:10,marginBottom:12}}>
        <span style={{fontSize:12,color:"#27ae60",fontWeight:600}}>✅ {t.recoveryKeyExists}</span>
        <button onClick={generate} style={{...btnC,fontSize:11,padding:"4px 12px"}}>{t.regenerateKey}</button>
      </div>}

      {!hasKey&&<button onClick={generate} disabled={generating} style={{...btnP,display:"flex",alignItems:"center",gap:6}}>
        <span>🔑</span>{generating?"...":t.generateKey}
      </button>}

      {showKey&&<div style={{background:"#fef9e7",border:"2px solid #c9a84c",borderRadius:10,padding:20,marginTop:12}}>
        <div style={{fontSize:12,fontWeight:600,color:"#7d6608",marginBottom:8}}>{t.recoveryKeyDesc}</div>
        <div style={{background:"#fff",border:"1px solid #e8d5a3",borderRadius:8,padding:"16px 20px",textAlign:"center",marginBottom:12}}>
          <div style={{fontFamily:"monospace",fontSize:22,fontWeight:700,color:"#0e3554",letterSpacing:2,userSelect:"all"}}>{showKey}</div>
        </div>
        <div style={{display:"flex",gap:8}}>
          <button onClick={()=>{navigator.clipboard?.writeText(showKey);}} style={{...btnC,fontSize:11,padding:"6px 14px"}}>📋 Copy</button>
          <button onClick={()=>setShowKey("")} style={{...btnP,fontSize:11,padding:"6px 14px",background:"#27ae60"}}>{t.iSavedIt}</button>
        </div>
      </div>}
    </div>
  );
}

// ─── AUDIT LOG PAGE ───────────────────────────────────────────────────
function AuditPage({t,auditLog}){
  return(
    <div>
      <h2 style={{fontSize:18,fontWeight:700,margin:"0 0 3px"}}>{t.auditLog}</h2>
      <p style={{fontSize:12,color:"var(--color-text-secondary)",margin:"0 0 16px"}}>🔒 {t.auditLog} — {auditLog.length} entries</p>
      <div style={crd}>
        {auditLog.length===0&&<div style={{textAlign:"center",padding:30,color:"var(--color-text-tertiary)"}}>{t.noAuditEntries}</div>}
        {auditLog.length>0&&<table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
          <thead><tr style={{borderBottom:"1px solid var(--color-border-tertiary)"}}>
            <th style={{textAlign:"start",padding:"8px",fontWeight:600,color:"var(--color-text-secondary)",width:"20%"}}>{t.auditTime}</th>
            <th style={{textAlign:"start",padding:"8px",fontWeight:600,color:"var(--color-text-secondary)",width:"15%"}}>{t.auditUser}</th>
            <th style={{textAlign:"start",padding:"8px",fontWeight:600,color:"var(--color-text-secondary)",width:"10%"}}>{t.role}</th>
            <th style={{textAlign:"start",padding:"8px",fontWeight:600,color:"var(--color-text-secondary)",width:"25%"}}>{t.auditAction}</th>
            <th style={{textAlign:"start",padding:"8px",fontWeight:600,color:"var(--color-text-secondary)",width:"30%"}}>{t.auditDetails}</th>
          </tr></thead>
          <tbody>{auditLog.map(e=>(
            <tr key={e.id} style={{borderBottom:"0.5px solid var(--color-border-tertiary)"}}>
              <td style={{padding:"6px 8px",fontSize:11,color:"var(--color-text-tertiary)"}}>{new Date(e.timestamp).toLocaleString()}</td>
              <td style={{padding:"6px 8px",fontWeight:500}}>{e.user}</td>
              <td style={{padding:"6px 8px"}}><span style={{fontSize:10,padding:"1px 6px",borderRadius:8,background:e.role==="admin"?"#e8d5a3":"#eaf2f8",color:e.role==="admin"?"#7d6119":"#1a5276"}}>{e.role||"-"}</span></td>
              <td style={{padding:"6px 8px",fontWeight:500,color:e.action.includes("failed")?"#e74c3c":e.action.includes("successful")?"#27ae60":"var(--color-text-primary)"}}>{e.action}</td>
              <td style={{padding:"6px 8px",fontSize:11,color:"var(--color-text-secondary)"}}>{e.details||"-"}</td>
            </tr>
          ))}</tbody>
        </table>}
      </div>
    </div>
  );
}

// ─── ACCESS PAGE ──────────────────────────────────────────────────────
function AccessPage({t}){
  return(
    <div>
      <h2 style={{fontSize:18,fontWeight:700,margin:"0 0 16px"}}>{t.access}</h2>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:14}}>
        {[{r:t.admin,c:"#c9a84c",p:t.adminPerms},{r:t.collegeRep,c:"#27ae60",p:t.collegePerms},{r:t.viewer,c:"#3498db",p:t.viewerPerms}].map(x=>(
          <div key={x.r} style={{...crd,background:"var(--color-background-secondary)"}}>
            <div style={{fontSize:14,fontWeight:700,color:x.c,marginBottom:8}}>{x.r}</div>
            <p style={{fontSize:12,color:"var(--color-text-secondary)",lineHeight:1.8}}>{x.p}</p>
          </div>
        ))}
      </div>
    </div>
  );
}

// ─── BACKUP PAGE ──────────────────────────────────────────────────────
function BackupPage({t,plans,inits,trainers,corr,reports,users,config,notes,savePlans,saveInits,saveTrainers,saveCorr,saveReports,saveUsers,saveConfig,saveNotes}){
  const [msg,sMsg]=useState("");
  const exportAll=()=>{const data={version:3,exportDate:new Date().toISOString(),plans,inits,trainers,corr,reports,users,config,notes};const blob=new Blob([JSON.stringify(data,null,2)],{type:"application/json"});downloadFile(blob,`CED_Backup_${new Date().toISOString().split("T")[0]}.json`);};
  const importData=e=>{const f=e.target.files[0];if(!f)return;const r=new FileReader();r.onload=async ev=>{try{const d=JSON.parse(ev.target.result);if(d.plans)await savePlans(d.plans);if(d.inits)await saveInits(d.inits);if(d.trainers)await saveTrainers(d.trainers);if(d.corr)await saveCorr(d.corr);if(d.reports)await saveReports(d.reports);if(d.users)await saveUsers(d.users);if(d.config)await saveConfig(d.config);if(d.notes)await saveNotes(d.notes);sMsg("✅ "+t.importData);}catch{sMsg("❌ Error");}};r.readAsText(f);};

  // Storage calculation
  const calcSize=(data)=>{try{return new Blob([JSON.stringify(data)]).size;}catch{return 0;}};
  const countFiles=(arr,key="fileData")=>arr.filter(x=>x[key]&&x[key].length>100).length;
  const countPlanFiles=()=>{let c=0;plans.forEach(p=>p.activities?.forEach(a=>{if(a.fileData&&a.fileData.length>100)c++;}));return c;};

  const LIMIT=5*1024*1024;
  const storageItems=[
    {key:t.yearlyPlans,data:plans,size:calcSize(plans),records:plans.reduce((s,p)=>s+p.activities.length,0),files:countPlanFiles(),save:savePlans,clearFn:()=>savePlans(plans.map(p=>({...p,activities:p.activities.map(a=>({...a,fileData:"",fileName:""}))})))},
    {key:t.initiatives,data:inits,size:calcSize(inits),records:inits.length,files:countFiles(inits),save:saveInits,clearFn:()=>saveInits(inits.map(x=>({...x,fileData:"",fileName:""})))},
    {key:t.trainers,data:trainers,size:calcSize(trainers),records:trainers.length,files:countFiles(trainers),save:saveTrainers,clearFn:()=>saveTrainers(trainers.map(x=>({...x,fileData:"",fileName:""})))},
    {key:t.correspondence,data:corr,size:calcSize(corr),records:corr.length,files:countFiles(corr),save:saveCorr,clearFn:()=>saveCorr(corr.map(x=>({...x,fileData:"",fileName:""})))},
    {key:t.reports,data:reports,size:calcSize(reports),records:reports.length,files:0},
    {key:t.stickyNotes||"Notes",data:notes,size:calcSize(notes),records:notes.length,files:0},
    {key:t.auditLog||"Audit",data:null,size:0,records:0,files:0,noCalc:true},
    {key:t.settings,data:config,size:calcSize(config),records:1,files:config.customLogo?1:0},
  ];
  const totalSize=storageItems.reduce((s,i)=>s+i.size,0);
  const fmtSize=(b)=>{if(b<1024)return b+" B";if(b<1024*1024)return (b/1024).toFixed(1)+" KB";return (b/(1024*1024)).toFixed(2)+" MB";};
  const pct=(b)=>Math.min(100,Math.round(b/LIMIT*100));

  return(
    <div>
      <h2 style={{fontSize:18,fontWeight:700,margin:"0 0 3px"}}>{t.backup}</h2>
      <p style={{fontSize:12,color:"var(--color-text-secondary)",margin:"0 0 16px"}}>{t.backupSubtitle}</p>

      {/* STORAGE MONITOR */}
      <div style={{...crd,marginBottom:16}}>
        <h3 style={{fontSize:14,fontWeight:600,margin:"0 0 4px"}}>📊 {t.storageManager}</h3>
        <p style={{fontSize:11,color:"var(--color-text-tertiary)",margin:"0 0 14px"}}>{t.storageLimit} · {t.storageTotal}: <b>{fmtSize(totalSize)}</b></p>

        {/* Overall status bar */}
        <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:16}}>
          <div style={{flex:1,height:22,background:"var(--color-background-tertiary)",borderRadius:11,overflow:"hidden",position:"relative"}}>
            <div style={{height:"100%",width:`${pct(totalSize)}%`,background:pct(totalSize)>80?"#e74c3c":pct(totalSize)>50?"#f39c12":"#27ae60",borderRadius:11,transition:"width 0.5s"}}/>
            <span style={{position:"absolute",top:"50%",left:"50%",transform:"translate(-50%,-50%)",fontSize:10,fontWeight:700,color:pct(totalSize)>40?"#fff":"var(--color-text-primary)"}}>{fmtSize(totalSize)} / 5 MB per key</span>
          </div>
        </div>

        {/* Per-category breakdown */}
        <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
          <thead><tr style={{borderBottom:"1px solid var(--color-border-tertiary)"}}>
            <th style={{textAlign:"start",padding:"8px",fontWeight:600,color:"var(--color-text-secondary)"}}>{t.storageKey}</th>
            <th style={{textAlign:"start",padding:"8px",fontWeight:600,color:"var(--color-text-secondary)"}}>{t.recordCount}</th>
            <th style={{textAlign:"start",padding:"8px",fontWeight:600,color:"var(--color-text-secondary)"}}>{t.withFiles}</th>
            <th style={{textAlign:"start",padding:"8px",fontWeight:600,color:"var(--color-text-secondary)"}}>{t.storageSize}</th>
            <th style={{textAlign:"start",padding:"8px",fontWeight:600,color:"var(--color-text-secondary)"}}>{t.storageUsage}</th>
            <th style={{padding:"8px"}}></th>
          </tr></thead>
          <tbody>{storageItems.map((it,i)=>{
            const p=pct(it.size);
            return(
            <tr key={i} style={{borderBottom:"0.5px solid var(--color-border-tertiary)"}}>
              <td style={{padding:"8px",fontWeight:500}}>{it.key}</td>
              <td style={{padding:"8px",color:"var(--color-text-secondary)"}}>{it.records}</td>
              <td style={{padding:"8px",color:it.files>0?"#e67e22":"var(--color-text-tertiary)"}}>{it.files>0?`📎 ${it.files}`:"-"}</td>
              <td style={{padding:"8px",fontWeight:600,color:p>80?"#e74c3c":p>50?"#f39c12":"var(--color-text-primary)"}}>{fmtSize(it.size)}</td>
              <td style={{padding:"8px",width:"25%"}}><div style={{height:8,background:"var(--color-background-tertiary)",borderRadius:4,overflow:"hidden"}}><div style={{height:"100%",width:`${p}%`,background:p>80?"#e74c3c":p>50?"#f39c12":"#27ae60",borderRadius:4}}/></div></td>
              <td style={{padding:"8px"}}>{it.files>0&&it.clearFn&&<button onClick={()=>{it.clearFn();sMsg("✅ "+t.storageFreed);}} style={{...btnC,padding:"3px 8px",fontSize:10,color:"#e67e22",borderColor:"#e67e22"}}>{t.clearAttachments}</button>}</td>
            </tr>);
          })}</tbody>
        </table>

        {/* Status message */}
        <div style={{marginTop:12,padding:10,borderRadius:8,fontSize:12,background:pct(totalSize)>80?"#fdecea":pct(totalSize)>50?"#fef9e7":"#d5f5e3",color:pct(totalSize)>80?"#922b21":pct(totalSize)>50?"#7d6608":"#196f3d"}}>
          {pct(totalSize)>80?`⚠️ ${t.storageDanger}`:pct(totalSize)>50?`⚠️ ${t.storageWarning}`:`✅ ${t.storageSafe}`}
        </div>
        <p style={{fontSize:11,color:"var(--color-text-tertiary)",marginTop:8,lineHeight:1.7}}>{t.clearAttachmentsDesc}</p>
      </div>

      {/* FILE OFFLOAD */}
      <div style={{...crd,marginBottom:16}}>
        <h3 style={{fontSize:14,fontWeight:600,margin:"0 0 4px"}}>💾 {t.offloadFiles}</h3>
        <p style={{fontSize:11,color:"var(--color-text-tertiary)",margin:"0 0 14px",lineHeight:1.7}}>{t.offloadDesc}</p>
        <OffloadPanel t={t} plans={plans} inits={inits} trainers={trainers} corr={corr} savePlans={savePlans} saveInits={saveInits} saveTrainers={saveTrainers} saveCorr={saveCorr}/>
      </div>

      {/* ARCHIVE SYSTEM */}
      <div style={{...crd,marginBottom:16}}>
        <h3 style={{fontSize:14,fontWeight:600,margin:"0 0 4px"}}>🗄️ {t.archiveManager}</h3>
        <p style={{fontSize:11,color:"var(--color-text-tertiary)",margin:"0 0 14px",lineHeight:1.7}}>{t.archiveDesc}</p>
        <ArchivePanel t={t} plans={plans} inits={inits} trainers={trainers} corr={corr} reports={reports}
          savePlans={savePlans} saveInits={saveInits} saveTrainers={saveTrainers} saveCorr={saveCorr} saveReports={saveReports}/>
      </div>

      {/* SELECTIVE EXPORT */}
      <div style={{...crd,marginBottom:16}}>
        <h3 style={{fontSize:14,fontWeight:600,margin:"0 0 8px"}}>📋 {t.selectiveExport}</h3>
        <SelectiveExport t={t} plans={plans} inits={inits} trainers={trainers} corr={corr} reports={reports} notes={notes} config={config} users={users}/>
      </div>

      {/* EXPORT / IMPORT */}
      <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(300px,1fr))",gap:14,marginBottom:16}}>
        <div style={crd}><h3 style={{fontSize:14,fontWeight:600,margin:"0 0 10px"}}>📤 {t.exportBackup}</h3><p style={{fontSize:12,color:"var(--color-text-secondary)",marginBottom:12}}>{t.downloadBackup}</p><button onClick={exportAll} style={{...btnP,display:"flex",alignItems:"center",gap:6}}><Ic n="download" s={15} c="#fff"/>{t.exportAll}</button></div>
        <div style={crd}><h3 style={{fontSize:14,fontWeight:600,margin:"0 0 10px"}}>📥 {t.importRestore}</h3><p style={{fontSize:12,color:"var(--color-text-secondary)",marginBottom:12}}>{t.importData}</p><input type="file" accept=".json" onChange={importData} style={{fontSize:12}}/>{msg&&<div style={{marginTop:8,fontSize:12}}>{msg}</div>}</div>
      </div>
      <div style={crd}><h3 style={{fontSize:14,fontWeight:600,margin:"0 0 8px"}}>{t.storageInfo}</h3><p style={{fontSize:12,color:"var(--color-text-secondary)",lineHeight:1.8}}>{t.storageDesc}</p><h3 style={{fontSize:14,fontWeight:600,margin:"16px 0 8px"}}>{t.modifyInfo}</h3><p style={{fontSize:12,color:"var(--color-text-secondary)",lineHeight:1.8}}>{t.modifyDesc}</p></div>
    </div>
  );
}

// ─── FILE OFFLOAD PANEL ───────────────────────────────────────────────
// ─── ARCHIVE PANEL ────────────────────────────────────────────────────
function ArchivePanel({t,plans,inits,trainers,corr,reports,savePlans,saveInits,saveTrainers,saveCorr,saveReports}){
  const [archives,setArchives]=useState([]);
  const [loadedArchive,setLoadedArchive]=useState(null);
  const [msg,setMsg]=useState("");
  const [archiveYear,setArchiveYear]=useState("2024-2025");

  useEffect(()=>{(async()=>{
    const keys=await sG("cedp-archive-index",[]);
    setArchives(keys);
  })();},[]);

  const doArchive=async()=>{
    const yearPlans=plans.filter(p=>String(p.year)===archiveYear);
    if(yearPlans.length===0&&inits.length===0){setMsg("⚠️ No data for "+archiveYear);return;}
    // Save stats snapshot before archiving
    const stats={totalActivities:yearPlans.reduce((s,p)=>s+p.activities.length,0),totalPlans:yearPlans.length,totalInits:inits.length,totalTrainers:trainers.length,totalCorr:corr.length,totalReports:reports.length,archivedAt:new Date().toISOString()};
    const archiveData={year:archiveYear,plans:yearPlans,inits,trainers,corr,reports,stats};
    const key="cedp-archive-"+archiveYear;
    await sS(key,archiveData);
    // Update index
    const idx=[...archives.filter(a=>a!==archiveYear),archiveYear];
    await sS("cedp-archive-index",idx);
    setArchives(idx);
    // Remove archived plans from active data
    const remainingPlans=plans.filter(p=>String(p.year)!==archiveYear);
    await savePlans(remainingPlans);
    setMsg("✅ "+t.archiveSaved+" ("+archiveYear+")");
  };

  const loadArch=async(year)=>{
    const data=await sG("cedp-archive-"+year,null);
    if(data){setLoadedArchive(data);setMsg("✅ "+t.archiveLoaded+" — "+year);}
  };

  const deleteArch=async(year)=>{
    try{localStorage.removeItem("cedp-archive-"+year);}catch{}
    const idx=archives.filter(a=>a!==year);
    await sS("cedp-archive-index",idx);
    setArchives(idx);
    if(loadedArchive?.year===year)setLoadedArchive(null);
    setMsg("🗑️ Deleted "+year);
  };

  const exportArch=(data)=>{
    const blob=new Blob([JSON.stringify(data,null,2)],{type:"application/json"});
    downloadFile(blob,`CED_Archive_${data.year}.json`);
  };

  return(
    <div>
      {/* Archive current year */}
      <div style={{display:"flex",gap:10,alignItems:"center",marginBottom:12,flexWrap:"wrap"}}>
        <select value={archiveYear} onChange={e=>setArchiveYear(e.target.value)} style={{...inp,width:"auto",minWidth:140}}>
          <option value="2024-2025">2024-2025</option>
          <option value="2025-2026">2025-2026</option>
          <option value="2026-2027">2026-2027</option>
        </select>
        <button onClick={doArchive} style={{...btnP,padding:"8px 16px",fontSize:12,display:"flex",alignItems:"center",gap:4}}>🗄️ {t.archiveYear}</button>
        <span style={{fontSize:11,color:"var(--color-text-tertiary)"}}>{t.activeData}: {plans.length} plans, {plans.reduce((s,p)=>s+p.activities.length,0)} activities</span>
      </div>
      {msg&&<div style={{fontSize:12,marginBottom:10,color:msg.includes("✅")?"#27ae60":msg.includes("⚠️")?"#e67e22":"#e74c3c"}}>{msg}</div>}

      {/* Archived years */}
      {archives.length>0&&<div>
        <div style={{fontSize:12,fontWeight:600,marginBottom:8}}>{t.archivedYears}:</div>
        {archives.map(yr=>(
          <div key={yr} style={{display:"flex",alignItems:"center",gap:8,padding:"8px 12px",background:"var(--color-background-secondary)",borderRadius:6,marginBottom:6}}>
            <span style={{fontSize:13}}>🗄️</span>
            <span style={{flex:1,fontSize:12,fontWeight:600}}>{yr}</span>
            <button onClick={()=>loadArch(yr)} style={{...btnC,padding:"3px 10px",fontSize:10}}>{t.loadArchive}</button>
            <button onClick={async()=>{const d=await sG("cedp-archive-"+yr,null);if(d)exportArch(d);}} style={{...btnC,padding:"3px 10px",fontSize:10}}>📤 Export</button>
            <button onClick={()=>deleteArch(yr)} style={{...btnC,padding:"3px 10px",fontSize:10,color:"#e74c3c",borderColor:"#e74c3c"}}>{t.deleteArchive}</button>
          </div>
        ))}
      </div>}

      {/* Loaded archive preview */}
      {loadedArchive&&<div style={{marginTop:12,padding:12,background:"#eaf2f8",borderRadius:8,border:"1px solid #2980b9"}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:8}}>
          <span style={{fontSize:13,fontWeight:700,color:"#1a5276"}}>📂 {loadedArchive.year}</span>
          <button onClick={()=>exportArch(loadedArchive)} style={{...btnP,padding:"4px 12px",fontSize:11}}>📤 Export full archive</button>
        </div>
        {loadedArchive.stats&&<div style={{fontSize:11,color:"#2c3e50",lineHeight:1.8}}>
          {t.plansData}: {loadedArchive.stats.totalPlans} · Activities: {loadedArchive.stats.totalActivities} · {t.initData}: {loadedArchive.stats.totalInits} · {t.trainersData}: {loadedArchive.stats.totalTrainers} · {t.corrData}: {loadedArchive.stats.totalCorr} · {t.reportsData}: {loadedArchive.stats.totalReports}
        </div>}
        <div style={{fontSize:10,color:"#5d6d7e",marginTop:4}}>{t.archiveNote}</div>
      </div>}
    </div>
  );
}

// ─── SELECTIVE EXPORT ─────────────────────────────────────────────────
function SelectiveExport({t,plans,inits,trainers,corr,reports,notes,config,users}){
  const [inc,setInc]=useState({plans:true,inits:true,trainers:true,corr:true,reports:true,notes:true,config:true,users:false});
  const [dateFrom,setDateFrom]=useState("");
  const [dateTo,setDateTo]=useState("");
  const [college,setCollege]=useState("All");
  const toggle=k=>setInc({...inc,[k]:!inc[k]});

  const doExport=()=>{
    const data={version:3,exportDate:new Date().toISOString(),exportType:"selective"};
    // Filter plans by college and date
    let filteredPlans=plans;
    if(college!=="All")filteredPlans=plans.filter(p=>p.college===college);
    if(dateFrom||dateTo){
      filteredPlans=filteredPlans.map(p=>({...p,activities:p.activities.filter(a=>{
        if(!a.date)return true;
        if(dateFrom&&a.date<dateFrom)return false;
        if(dateTo&&a.date>dateTo)return false;
        return true;
      })}));
    }
    if(inc.plans)data.plans=filteredPlans;
    if(inc.inits)data.inits=inits;
    if(inc.trainers)data.trainers=trainers;
    if(inc.corr){
      let fc=corr;
      if(dateFrom||dateTo)fc=corr.filter(c=>{if(!c.date)return true;if(dateFrom&&c.date<dateFrom)return false;if(dateTo&&c.date>dateTo)return false;return true;});
      data.corr=fc;
    }
    if(inc.reports)data.reports=reports;
    if(inc.notes)data.notes=notes;
    if(inc.config)data.config=config;
    if(inc.users)data.users=users;
    const blob=new Blob([JSON.stringify(data,null,2)],{type:"application/json"});
    const label=college!=="All"?college.replace(/College of |Department of /g,""):"All";
    downloadFile(blob,`CED_Export_${label}_${dateFrom||"start"}_to_${dateTo||"end"}.json`);
  };

  const colleges=[...new Set(plans.map(p=>p.college))];
  const items=[
    {k:"plans",l:t.plansData,c:plans.length},{k:"inits",l:t.initData,c:inits.length},
    {k:"trainers",l:t.trainersData,c:trainers.length},{k:"corr",l:t.corrData,c:corr.length},
    {k:"reports",l:t.reportsData,c:reports.length},{k:"notes",l:t.stickyNotes||"Notes",c:notes.length},
  ];

  return(
    <div>
      {/* Category checkboxes */}
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:6,marginBottom:12}}>
        {items.map(it=>(
          <label key={it.k} style={{display:"flex",alignItems:"center",gap:6,fontSize:12,cursor:"pointer",padding:"6px 10px",borderRadius:6,background:inc[it.k]?"#eaf2f8":"var(--color-background-secondary)",border:`1px solid ${inc[it.k]?"#2980b9":"var(--color-border-tertiary)"}`}}>
            <input type="checkbox" checked={inc[it.k]} onChange={()=>toggle(it.k)} style={{accentColor:"#1a5276"}}/>
            {it.l} ({it.c})
          </label>
        ))}
      </div>
      {/* Filters */}
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:10,marginBottom:12}}>
        <div><label style={lbl}>{t.college}</label><select value={college} onChange={e=>setCollege(e.target.value)} style={inp}><option value="All">{t.all}</option>{colleges.map(c=><option key={c}>{c}</option>)}</select></div>
        <div><label style={lbl}>{t.dateFrom}</label><input type="date" value={dateFrom} onChange={e=>setDateFrom(e.target.value)} style={inp}/></div>
        <div><label style={lbl}>{t.dateTo}</label><input type="date" value={dateTo} onChange={e=>setDateTo(e.target.value)} style={inp}/></div>
      </div>
      <button onClick={doExport} style={{...btnP,display:"flex",alignItems:"center",gap:6}}><Ic n="download" s={15} c="#fff"/>{t.exportSelected}</button>
    </div>
  );
}

function OffloadPanel({t,plans,inits,trainers,corr,savePlans,saveInits,saveTrainers,saveCorr}){
  const [downloading,setDownloading]=useState(false);
  const [result,setResult]=useState("");
  const [selected,setSelected]=useState({});

  const allFiles=useMemo(()=>{
    const files=[];
    plans.forEach(p=>p.activities?.forEach(a=>{if(a.fileData&&a.fileData.length>100)files.push({uid:`plan-${p.id}-${a.id}`,category:t.yearlyPlans,label:`${p.college} — ${a.title}`,name:a.fileName||"activity_file",data:a.fileData,source:"plans",planId:p.id,actId:a.id});}));
    inits.forEach(i=>{if(i.fileData&&i.fileData.length>100)files.push({uid:`init-${i.id}`,category:t.initiatives,label:i.title,name:i.fileName||"initiative_file",data:i.fileData,source:"inits",id:i.id});});
    trainers.forEach(x=>{if(x.fileData&&x.fileData.length>100)files.push({uid:`trainer-${x.id}`,category:t.trainers,label:x.name,name:x.fileName||"trainer_file",data:x.fileData,source:"trainers",id:x.id});});
    corr.forEach(c=>{if(c.fileData&&c.fileData.length>100)files.push({uid:`corr-${c.id}`,category:t.correspondence,label:`${c.refNumber||""} — ${c.subject}`,name:c.fileName||"correspondence_file",data:c.fileData,source:"corr",id:c.id});});
    return files;
  },[plans,inits,trainers,corr,t]);

  // Initialize selection when files change
  useEffect(()=>{
    const sel={};allFiles.forEach(f=>{sel[f.uid]=selected[f.uid]??true;});
    setSelected(sel);
  },[allFiles.length]);

  const selCount=Object.values(selected).filter(Boolean).length;
  const selAll=()=>{const s={};allFiles.forEach(f=>{s[f.uid]=true;});setSelected(s);};
  const selNone=()=>{const s={};allFiles.forEach(f=>{s[f.uid]=false;});setSelected(s);};

  const downloadSelected=async()=>{
    const toDownload=allFiles.filter(f=>selected[f.uid]);
    if(toDownload.length===0)return;
    setDownloading(true);setResult("");
    let count=0;
    for(const f of toDownload){
      try{
        downloadFile(f.data, f.name);
        count++;
        // Longer delay to prevent browser blocking
        await new Promise(r=>setTimeout(r,1000));
      }catch(e){console.error("Download failed:",f.name,e);}
    }
    setResult(`✅ ${count} ${t.filesDownloaded}`);
    setDownloading(false);
  };

  const clearSelected=()=>{
    const selUids=new Set(allFiles.filter(f=>selected[f.uid]).map(f=>f.uid));
    const selPlanActs=new Set();const selInitIds=new Set();const selTrainerIds=new Set();const selCorrIds=new Set();
    allFiles.filter(f=>selUids.has(f.uid)).forEach(f=>{
      if(f.source==="plans")selPlanActs.add(`${f.planId}|${f.actId}`);
      else if(f.source==="inits")selInitIds.add(f.id);
      else if(f.source==="trainers")selTrainerIds.add(f.id);
      else if(f.source==="corr")selCorrIds.add(f.id);
    });
    if(selPlanActs.size>0)savePlans(plans.map(p=>({...p,activities:p.activities.map(a=>selPlanActs.has(`${p.id}|${a.id}`)?{...a,fileData:"",fileName:"",notes:(a.notes||"")+(a.notes?" | ":"")+"📁 "+a.fileName}:a)})));
    if(selInitIds.size>0)saveInits(inits.map(i=>selInitIds.has(i.id)?{...i,fileData:"",fileName:"",notes:(i.notes||"")+(i.notes?" | ":"")+"📁 "+i.fileName}:i));
    if(selTrainerIds.size>0)saveTrainers(trainers.map(x=>selTrainerIds.has(x.id)?{...x,fileData:"",fileName:"",notes:(x.notes||"")+(x.notes?" | ":"")+"📁 "+x.fileName}:x));
    if(selCorrIds.size>0)saveCorr(corr.map(c=>selCorrIds.has(c.id)?{...c,fileData:"",fileName:"",notes:(c.notes||"")+(c.notes?" | ":"")+"📁 "+c.fileName}:c));
    setResult("✅ "+selCount+" files cleared from storage");
  };

  if(allFiles.length===0)return <div style={{fontSize:12,color:"var(--color-text-tertiary)",padding:"8px 0"}}>{t.noFilesStored}</div>;

  return(
    <div>
      {/* Select controls */}
      <div style={{display:"flex",gap:8,marginBottom:10,alignItems:"center"}}>
        <button onClick={selAll} style={{...btnC,padding:"3px 10px",fontSize:10}}>{t.selectAll}</button>
        <button onClick={selNone} style={{...btnC,padding:"3px 10px",fontSize:10}}>{t.deselectAll}</button>
        <span style={{fontSize:11,color:"var(--color-text-tertiary)"}}>{selCount}/{allFiles.length} selected</span>
      </div>

      {/* File list with checkboxes */}
      <div style={{maxHeight:220,overflowY:"auto",marginBottom:12,border:"1px solid var(--color-border-tertiary)",borderRadius:8}}>
        {allFiles.map((f,i)=>(
          <div key={f.uid} style={{display:"flex",alignItems:"center",gap:8,padding:"6px 10px",borderBottom:i<allFiles.length-1?"0.5px solid var(--color-border-tertiary)":"none",fontSize:11,background:selected[f.uid]?"#eaf2f8":"transparent"}}>
            <input type="checkbox" checked={!!selected[f.uid]} onChange={()=>setSelected({...selected,[f.uid]:!selected[f.uid]})} style={{accentColor:"#1a5276",flexShrink:0}}/>
            <span style={{fontSize:12}}>📎</span>
            <div style={{flex:1,minWidth:0}}>
              <div style={{fontWeight:500,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{f.name}</div>
              <div style={{color:"var(--color-text-tertiary)",fontSize:10}}>{f.category} — {f.label}</div>
            </div>
            <span style={{fontSize:10,color:"var(--color-text-tertiary)",flexShrink:0}}>{(f.data.length/1024).toFixed(0)}KB</span>
            <button onClick={()=>downloadFile(f.data,f.name)} style={{...btnI,padding:3,flexShrink:0}}><Ic n="download" s={12} c="#2980b9"/></button>
          </div>
        ))}
      </div>

      {/* Action buttons — separate download and clear */}
      <div style={{display:"flex",gap:10,flexWrap:"wrap"}}>
        <button onClick={downloadSelected} disabled={downloading||selCount===0} style={{...btnP,display:"flex",alignItems:"center",gap:6,opacity:(downloading||selCount===0)?0.5:1}}>
          <Ic n="download" s={15} c="#fff"/>{downloading?t.downloading:`Download selected (${selCount})`}
        </button>
        <button onClick={clearSelected} disabled={selCount===0} style={{...btnP,display:"flex",alignItems:"center",gap:6,background:"#e74c3c",opacity:selCount===0?0.5:1}}>
          🗑️ Clear selected from storage ({selCount})
        </button>
      </div>
      {result&&<div style={{marginTop:10,fontSize:12,color:"#27ae60",fontWeight:500}}>{result}</div>}

      <div style={{marginTop:12,padding:10,background:"#fef9e7",borderRadius:8,fontSize:11,color:"#7d6608",lineHeight:1.7}}>
        💡 Download files first, then clear. Each file has its own download button. Clear only removes selected files.
      </div>
    </div>
  );
}

// ─── UNIVERSAL PRINT PREVIEW ──────────────────────────────────────────
function showPrintPreview(html) {
  const old=document.getElementById('print-overlay');if(old)old.remove();
  const oldBar=document.getElementById('print-bar');if(oldBar)oldBar.remove();
  const overlay=document.createElement('div');
  overlay.id='print-overlay';
  overlay.style.cssText='position:fixed;top:0;left:0;width:100%;height:100%;z-index:99999;background:#fff;overflow:auto;';
  const style=document.createElement('style');
  style.textContent=`@media print{body>*:not(#print-overlay){display:none!important;}#print-overlay{position:static!important;overflow:visible!important;}#print-bar{display:none!important;}}`;
  overlay.appendChild(style);
  const content=document.createElement('div');
  content.innerHTML=html.replace(/<html[^>]*>|<\/html>|<head>[\s\S]*?<\/head>|<body[^>]*>|<\/body>/gi,'');
  const styleMatch=html.match(/<style>([\s\S]*?)<\/style>/);
  if(styleMatch){const s=document.createElement('style');s.textContent=styleMatch[1];content.prepend(s);}
  if(html.includes('dir="rtl"'))content.dir='rtl';
  overlay.appendChild(content);
  document.body.appendChild(overlay);
  const bar=document.createElement('div');
  bar.id='print-bar';
  bar.style.cssText='position:fixed;top:10px;right:10px;z-index:100001;display:flex;gap:8px;flex-wrap:wrap;';
  const mkBtn=(text,bg,fn)=>{const b=document.createElement('button');b.textContent=text;b.style.cssText=`padding:10px 18px;background:${bg};color:#fff;border:none;border-radius:6px;cursor:pointer;font-size:13px;font-weight:600;box-shadow:0 2px 8px rgba(0,0,0,0.3);`;b.onclick=fn;return b;};
  bar.appendChild(mkBtn('Print (Ctrl+P)','#1a5276',()=>{try{window.print();}catch(e){alert("Use Ctrl+P to print");}}));
  bar.appendChild(mkBtn('Save report','#27ae60',()=>{
    const blob=new Blob(["<!DOCTYPE html>"+html],{type:"text/html;charset=utf-8"});
    const url=URL.createObjectURL(blob);
    const a=document.createElement("a");a.href=url;a.download="report_"+new Date().toISOString().split("T")[0]+".html";
    document.body.appendChild(a);a.click();document.body.removeChild(a);URL.revokeObjectURL(url);
  }));
  bar.appendChild(mkBtn('Close','#e74c3c',()=>{overlay.remove();bar.remove();}));
  document.body.appendChild(bar);
}

// ─── DOWNLOAD HELPER (works in sandbox) ────────────────────────────
function downloadFile(dataOrBlob, filename) {
  try {
    let url;
    if (typeof dataOrBlob === 'string') {
      url = dataOrBlob;
    } else {
      url = URL.createObjectURL(dataOrBlob);
    }
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    a.style.display = 'none';
    document.body.appendChild(a);
    a.click();
    setTimeout(() => { document.body.removeChild(a); if (typeof dataOrBlob !== 'string') URL.revokeObjectURL(url); }, 500);
  } catch (e) {
    console.error('Download failed:', e);
    // Fallback: open in new tab
    try {
      const w = window.open();
      if (w) { w.document.write('<p>Right-click and Save As: ' + filename + '</p>'); w.document.close(); }
    } catch {}
  }
}

// ─── IMAGE COMPRESSION ────────────────────────────────────────────────
function compressImage(dataUrl, maxWidth=1200, quality=0.6) {
  return new Promise((resolve) => {
    const img = new Image();
    img.onload = () => {
      let w = img.width, h = img.height;
      if (w > maxWidth) { h = Math.round(h * maxWidth / w); w = maxWidth; }
      const canvas = document.createElement('canvas');
      canvas.width = w; canvas.height = h;
      const ctx = canvas.getContext('2d');
      ctx.drawImage(img, 0, 0, w, h);
      const compressed = canvas.toDataURL('image/jpeg', quality);
      console.log(`Image compressed: ${(dataUrl.length/1024).toFixed(0)}KB → ${(compressed.length/1024).toFixed(0)}KB`);
      resolve(compressed);
    };
    img.onerror = () => resolve(dataUrl); // fallback to original
    img.src = dataUrl;
  });
}

async function processFileForStorage(file) {
  const dataUrl = await new Promise((res, rej) => {
    const r = new FileReader(); r.onload = () => res(r.result); r.onerror = () => rej(); r.readAsDataURL(file);
  });
  // Compress images, keep PDFs/text as-is
  if (file.type && file.type.startsWith('image/')) {
    const compressed = await compressImage(dataUrl);
    return { dataUrl: compressed, originalDataUrl: dataUrl };
  }
  return { dataUrl, originalDataUrl: dataUrl };
}

// ─── EXPORT ENGINE ────────────────────────────────────────────────────
function buildReportHTML(a, opts, t, logoSrc, pName, pSub) {
  const isRtl = t.dir === "rtl";
  let html = `<html dir="${isRtl?'rtl':'ltr'}"><head><meta charset="utf-8"><title>${a.title}</title><style>
    @import url('https://fonts.googleapis.com/css2?family=Noto+Sans+Arabic:wght@400;700&display=swap');
    body{font-family:'Noto Sans Arabic','Noto Sans',sans-serif;padding:40px 50px;color:#2c3e50;font-size:14px;line-height:1.8;}
    .hdr{display:flex;align-items:center;gap:16px;border-bottom:3px solid #1a5276;padding-bottom:16px;margin-bottom:24px;}
    .hdr img{width:70px;height:70px;border-radius:50%;object-fit:contain;background:#0d3249;border:2px solid #c9a84c;}
    .hdr-txt h1{margin:0;font-size:20px;color:#1a5276;} .hdr-txt p{margin:2px 0;font-size:12px;color:#6c7a89;}
    h2{color:#1a5276;font-size:18px;border-bottom:1px solid #ddd;padding-bottom:8px;margin:24px 0 16px;}
    .row{display:flex;gap:12px;margin:6px 0;} .lbl{font-weight:700;min-width:180px;color:#5d6d7e;} .val{flex:1;}
    .badge{display:inline-block;padding:2px 10px;border-radius:12px;font-size:11px;font-weight:600;margin:2px 4px;}
    .sect{margin:20px 0;} table{width:100%;border-collapse:collapse;margin:10px 0;} th,td{border:1px solid #ddd;padding:8px 12px;text-align:${isRtl?'right':'left'};font-size:12px;} th{background:#f4f6f7;font-weight:700;}
    .footer{margin-top:30px;padding-top:12px;border-top:1px solid #ddd;font-size:11px;color:#999;}
  </style></head><body>`;
  if (opts.header) html += `<div class="hdr"><img src="${logoSrc}"/><div class="hdr-txt"><h1>${pName}</h1><p>${pSub}</p><p>${t.activityReport}</p></div></div>`;
  html += `<h2>${a.title}</h2>`;
  if (opts.details) {
    const fields = [{l:t.type,v:a.type},{l:t.status,v:a.status},{l:t.source,v:a.source},{l:t.delivery,v:a.delivery},{l:t.target,v:a.target},{l:t.presenter,v:a.presenter},{l:t.venue,v:a.venue},{l:t.duration,v:a.duration},{l:t.department,v:a.department}];
    if (opts.timeline) fields.unshift({l:t.date,v:a.date});
    if (opts.attendees) fields.push({l:t.attendees,v:String(a.attendees||0)});
    fields.forEach(f => { if(f.v) html += `<div class="row"><span class="lbl">${f.l}:</span><span class="val">${f.v}</span></div>`; });
  }
  if (opts.objectives && (a.objectives || a.notes)) {
    if (a.objectives) html += `<div class="row"><span class="lbl">${t.objectives}:</span><span class="val">${a.objectives}</span></div>`;
    if (a.notes) html += `<div class="row"><span class="lbl">${t.notes}:</span><span class="val">${a.notes}</span></div>`;
  }
  if (opts.media && a.mediaStatus && a.mediaStatus !== "Not contacted") {
    html += `<h2>📺 ${t.mediaCoverage}</h2>`;
    html += `<div class="row"><span class="lbl">${t.mediaStatus}:</span><span class="val">${a.mediaStatus}</span></div>`;
    if (a.mediaChannel) html += `<div class="row"><span class="lbl">${t.mediaChannel}:</span><span class="val">${a.mediaChannel}</span></div>`;
    if (a.mediaContactedDate) html += `<div class="row"><span class="lbl">${t.mediaContactedDate}:</span><span class="val">${a.mediaContactedDate}</span></div>`;
    if (a.mediaPublishedDate) html += `<div class="row"><span class="lbl">${t.mediaPublishedDate}:</span><span class="val">${a.mediaPublishedDate}</span></div>`;
    if (a.mediaCoverageNotes) html += `<div class="row"><span class="lbl">${t.mediaCoverageNotes}:</span><span class="val">${a.mediaCoverageNotes}</span></div>`;
  }
  if (opts.customFields && a.customFields && a.customFields.length > 0) {
    html += `<h2>📋 ${t.extractedFields}</h2>`;
    a.customFields.filter(f=>f.name).forEach(f => { html += `<div class="row"><span class="lbl">${f.name}:</span><span class="val">${f.value}</span></div>`; });
  }
  if (opts.participants && a.participants) {
    html += `<h2>${t.participantNames}</h2><p>${a.participants}</p>`;
  }
  html += `<div class="footer">${t.preparedBy}: ${pName} · ${new Date().toLocaleDateString()}</div></body></html>`;
  return html;
}

function exportAsPrint(a, opts, t, logoSrc, pName, pSub, setModal) {
  const html = buildReportHTML(a, opts, t, logoSrc, pName, pSub);
  showPrintPreview(html);
}

function exportBulkExcel(activities, t) {
  const data = activities.map(a => {
    const row = {};
    row[t.title] = a.title; row[t.type] = a.type; row[t.date] = a.date;
    row[t.status] = a.status; row[t.source] = a.source; row[t.presenter] = a.presenter || "";
    row[t.delivery] = a.delivery || ""; row[t.target] = a.target || "";
    row[t.venue] = a.venue || ""; row[t.duration] = a.duration || "";
    row[t.attendees] = a.attendees || 0; row[t.college] = a.college || "";
    row[t.objectives] = a.objectives || ""; row[t.notes] = a.notes || ""; row[t.department||"Department"] = a.department || "";
    row[t.mediaStatus] = a.mediaStatus || ""; row[t.mediaChannel] = a.mediaChannel || "";
    if (a.customFields) a.customFields.filter(f=>f.name).forEach(f => { row[f.name] = f.value; });
    return row;
  });
  const ws = XLSX.utils.json_to_sheet(data);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Activities");
  XLSX.writeFile(wb, `Activities_${new Date().toISOString().split('T')[0]}.xlsx`);
}

function exportSingleExcel(a, t) {
  const rows = [];
  const add = (l,v) => { if(v) rows.push({[t.fieldName||"Field"]:l,[t.fieldValue||"Value"]:v}); };
  add(t.title,a.title); add(t.type,a.type); add(t.date,a.date); add(t.status,a.status);
  add(t.source,a.source); add(t.presenter,a.presenter); add(t.delivery,a.delivery);
  add(t.target,a.target); add(t.venue,a.venue); add(t.duration,a.duration); add(t.department,a.department);
  add(t.attendees,String(a.attendees||0)); add(t.objectives,a.objectives); add(t.notes,a.notes);
  add(t.mediaStatus,a.mediaStatus); add(t.mediaChannel,a.mediaChannel);
  add(t.mediaContactedDate,a.mediaContactedDate); add(t.mediaPublishedDate,a.mediaPublishedDate);
  add(t.mediaCoverageNotes,a.mediaCoverageNotes);
  if (a.participants) add(t.participantNames,a.participants);
  if (a.customFields) a.customFields.filter(f=>f.name).forEach(f=>add(f.name,f.value));
  const ws = XLSX.utils.json_to_sheet(rows);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Report");
  XLSX.writeFile(wb, `${a.title.replace(/[^a-zA-Z0-9\u0600-\u06FF ]/g,'_')}_report.xlsx`);
}

function exportTrainersExcel(trainers, t) {
  const data = trainers.map(x => {
    const row = {};
    row[t.name] = x.name; row[t.dob||"DOB"] = x.dob; row[t.gender] = x.gender;
    row[t.phone] = x.phone; row[t.email] = x.email; row[t.workplace] = x.workplace;
    row[t.degree] = x.degree; row[t.generalSpec] = x.generalSpec; row[t.specificSpec] = x.specificSpec;
    row[t.internal+"/"+t.external] = x.internal ? t.internal : t.external;
    row[t.college] = x.college || ""; row[t.notes] = x.notes || "";
    if (x.linkedActivities) row["Linked activities"] = x.linkedActivities.join(", ");
    if (x.customFields) x.customFields.filter(f=>f.name).forEach(f => { row[f.name] = f.value; });
    return row;
  });
  const ws = XLSX.utils.json_to_sheet(data);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Trainers");
  XLSX.writeFile(wb, `Trainers_${new Date().toISOString().split('T')[0]}.xlsx`);
}

function printTrainers(trainers, t, logoSrc, pName) {
  const isRtl = t.dir === "rtl";
  let html = `<html dir="${isRtl?'rtl':'ltr'}"><head><meta charset="utf-8"><style>
    @import url('https://fonts.googleapis.com/css2?family=Noto+Sans+Arabic:wght@400;700&display=swap');
    body{font-family:'Noto Sans Arabic','Noto Sans',sans-serif;padding:30px;font-size:13px;}
    table{width:100%;border-collapse:collapse;margin:16px 0;} th,td{border:1px solid #ddd;padding:6px 10px;text-align:${isRtl?"right":"left"};font-size:11px;} th{background:#f4f6f7;font-weight:700;}
  </style></head><body>`;
  if (logoSrc) html += `<div style="display:flex;align-items:center;gap:14px;border-bottom:3px solid #1a5276;padding-bottom:14px;margin-bottom:20px;"><img src="${logoSrc}" width="60" style="border-radius:50%;background:#0d3249;"/><div><h1 style="margin:0;font-size:18px;color:#1a5276;">${pName||""}</h1><p style="margin:2px 0;font-size:11px;color:#6c7a89;">${t.trainers} — ${trainers.length}</p></div></div>`;
  html += `<h2>${t.trainers} (${trainers.length})</h2>`;
  html += `<table><thead><tr><th>#</th><th>${t.name}</th><th>${t.degree}</th><th>${t.specificSpec}</th><th>${t.workplace}</th><th>${t.college}</th><th>${t.phone}</th><th>${t.email}</th></tr></thead><tbody>`;
  trainers.forEach((x, i) => {
    html += `<tr><td>${i+1}</td><td>${x.name}</td><td>${x.degree}</td><td>${x.specificSpec||""}</td><td>${x.workplace||""}</td><td>${(x.college||"").replace("College of ","").replace("Department of ","")}</td><td>${x.phone||""}</td><td>${x.email||""}</td></tr>`;
  });
  html += `</tbody></table><p style="color:#999;font-size:11px;">${new Date().toLocaleDateString()}</p></body></html>`;
  showPrintPreview(html);
}

function exportReportsExcel(reports, t) {
  const data = reports.map(r => {
    const row = {};
    row[t.title] = r.activityTitle; row[t.date] = r.activityDate; row[t.type||"Type"] = r.activityType || "";
    row[t.college] = r.college || ""; row[t.presenter] = r.presenter || "";
    row[t.totalParticipants] = r.totalParticipants || 0;
    row[t.participantNames] = r.participantNames || "";
    row[t.summary] = r.summary || "";
    row[t.mediaStatus] = r.mediaStatus || "Not contacted";
    row[t.mediaChannel] = r.mediaChannel || "";
    row[t.mediaContactedDate] = r.mediaContactedDate || "";
    row[t.mediaPublishedDate] = r.mediaPublishedDate || "";
    row[t.mediaLinks] = r.mediaLinks || "";
    row[t.mediaCoverageNotes] = r.mediaCoverageNotes || "";
    row[t.confirmed] = r.confirmed ? "✅" : "❌";
    row["Filed"] = r.createdDate || "";
    return row;
  });
  const ws = XLSX.utils.json_to_sheet(data);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Reports");
  XLSX.writeFile(wb, `Event_Reports_${new Date().toISOString().split('T')[0]}.xlsx`);
}

function printReports(reports, t, logoSrc, pName) {
  const isRtl = t.dir === "rtl";
  let html = `<html dir="${isRtl?'rtl':'ltr'}"><head><meta charset="utf-8"><style>
    @import url('https://fonts.googleapis.com/css2?family=Noto+Sans+Arabic:wght@400;700&display=swap');
    body{font-family:'Noto Sans Arabic','Noto Sans',sans-serif;padding:30px;font-size:13px;}
    table{width:100%;border-collapse:collapse;margin:16px 0;} th,td{border:1px solid #ddd;padding:6px 10px;text-align:${isRtl?"right":"left"};font-size:11px;} th{background:#f4f6f7;font-weight:700;}
    .rpt{page-break-inside:avoid;margin:20px 0;padding:16px;border:1px solid #ddd;border-radius:8px;}
  </style></head><body>`;
  if (logoSrc) html += `<div style="display:flex;align-items:center;gap:14px;border-bottom:3px solid #1a5276;padding-bottom:14px;margin-bottom:20px;"><img src="${logoSrc}" width="60" style="border-radius:50%;background:#0d3249;"/><div><h1 style="margin:0;font-size:18px;color:#1a5276;">${pName||""}</h1><p style="margin:2px 0;font-size:11px;color:#6c7a89;">${t.reports} — ${reports.length}</p></div></div>`;
  html += `<h2>${t.reports} (${reports.length})</h2>`;
  // Summary table
  html += `<table><thead><tr><th>#</th><th>${t.title}</th><th>${t.date}</th><th>${t.college}</th><th>${t.presenter}</th><th>${t.totalParticipants}</th><th>${t.mediaStatus}</th><th>${t.mediaChannel}</th><th>${t.confirmed}</th></tr></thead><tbody>`;
  reports.forEach((r, i) => {
    html += `<tr><td>${i+1}</td><td>${r.activityTitle}</td><td>${r.activityDate||""}</td><td>${(r.college||"").replace("College of ","").replace("Department of ","")}</td><td>${r.presenter||""}</td><td>${r.totalParticipants||0}</td><td>${r.mediaStatus||"Not contacted"}</td><td>${r.mediaChannel||""}</td><td>${r.confirmed?"✅":"❌"}</td></tr>`;
  });
  html += `</tbody></table>`;
  // Detailed section
  reports.forEach((r, i) => {
    html += `<div class="rpt"><h3>${i+1}. ${r.activityTitle}</h3>`;
    html += `<p><b>${t.date}:</b> ${r.activityDate||""} · <b>${t.college}:</b> ${r.college||""} · <b>${t.presenter}:</b> ${r.presenter||""} · <b>${t.totalParticipants}:</b> ${r.totalParticipants||0}</p>`;
    if (r.summary) html += `<p><b>${t.summary}:</b> ${r.summary}</p>`;
    if (r.participantNames) html += `<p><b>${t.participantNames}:</b> ${r.participantNames}</p>`;
    html += `</div>`;
  });
  html += `<p style="color:#999;font-size:11px;">${new Date().toLocaleDateString()}</p></body></html>`;
  showPrintPreview(html);
}

function ExportModal({t, activity, activities, logoSrc, pName, pSub, onClose, isBulk}) {
  const [opts, setOpts] = useState({header:true,details:true,timeline:true,objectives:true,media:true,customFields:true,participants:true,attendees:true});
  const [fCollege,setFCollege]=useState("All");
  const [fStatus,setFStatus]=useState("All");
  const [fDateFrom,setFDateFrom]=useState("");
  const [fDateTo,setFDateTo]=useState("");
  const [format,setFormat]=useState("");
  const toggle = k => setOpts({...opts,[k]:!opts[k]});
  const allOn = () => setOpts({header:true,details:true,timeline:true,objectives:true,media:true,customFields:true,participants:true,attendees:true});
  const allOff = () => setOpts({header:false,details:false,timeline:false,objectives:false,media:false,customFields:false,participants:false,attendees:false});

  const optItems = [
    {k:"header",l:t.incHeader},{k:"details",l:t.incDetails},{k:"timeline",l:t.incTimeline},
    {k:"objectives",l:t.incObjectives},{k:"attendees",l:t.incAttendees},
    {k:"media",l:t.incMedia},{k:"customFields",l:t.incCustomFields},{k:"participants",l:t.incParticipants},
  ];

  // Filter activities for bulk
  const filtered = isBulk ? (activities||[]).filter(a=>{
    if(fCollege!=="All"&&a.college!==fCollege)return false;
    if(fStatus!=="All"&&a.status!==fStatus)return false;
    if(fDateFrom&&a.date&&a.date<fDateFrom)return false;
    if(fDateTo&&a.date&&a.date>fDateTo)return false;
    return true;
  }) : [];

  const colleges = isBulk ? [...new Set((activities||[]).map(a=>a.college).filter(Boolean))] : [];

  const doExport=(fmt)=>{
    if(isBulk){
      if(fmt==="excel") exportBulkExcel(filtered,t);
      else if(fmt==="print"&&filtered.length>0){
        let html=`<html dir="${t.dir}"><head><meta charset="utf-8"><style>body{font-family:'Noto Sans Arabic','Noto Sans',sans-serif;padding:30px;font-size:13px;} table{width:100%;border-collapse:collapse;margin:16px 0;} th,td{border:1px solid #ddd;padding:6px 10px;text-align:${t.dir==="rtl"?"right":"left"};font-size:11px;} th{background:#f4f6f7;}</style></head><body>`;
        if(opts.header)html+=`<div style="display:flex;align-items:center;gap:14px;border-bottom:3px solid #1a5276;padding-bottom:14px;margin-bottom:20px;"><img src="${logoSrc||LOGO}" width="60" style="border-radius:50%;background:#0d3249;"/><div><h1 style="margin:0;font-size:18px;color:#1a5276;">${pName}</h1><p style="margin:2px 0;font-size:11px;color:#6c7a89;">${pSub}</p></div></div>`;
        html+=`<h2>${t.bulkReport} — ${filtered.length} ${t.results}</h2>`;
        html+=`<table><thead><tr><th>#</th><th>${t.title}</th><th>${t.type}</th><th>${t.date}</th><th>${t.status}</th><th>${t.presenter}</th><th>${t.college}</th>`;
        if(opts.attendees)html+=`<th>${t.attendees}</th>`;
        if(opts.media)html+=`<th>${t.mediaStatus}</th>`;
        html+=`</tr></thead><tbody>`;
        filtered.forEach((a,i)=>{
          html+=`<tr><td>${i+1}</td><td>${a.title}</td><td>${a.type}</td><td>${a.date||""}</td><td>${a.status}</td><td>${a.presenter||""}</td><td>${a.college||""}</td>`;
          if(opts.attendees)html+=`<td>${a.attendees||0}</td>`;
          if(opts.media)html+=`<td>${a.mediaStatus||""}</td>`;
          html+=`</tr>`;
        });
        html+=`</tbody></table><p style="color:#999;font-size:11px;">${t.preparedBy}: ${pName} · ${new Date().toLocaleDateString()}</p></body></html>`;
        showPrintPreview(html);
      }
    } else {
      if(fmt==="excel") exportSingleExcel(activity,t);
      else if(fmt==="print") { const html=buildReportHTML(activity,opts,t,logoSrc||LOGO,pName,pSub); showPrintPreview(html); }
    }
    onClose();
  };

  return (
    <div>
      <h3 style={{fontSize:16,fontWeight:700,margin:"0 0 14px",color:"#1a5276"}}>{t.exportOptions}</h3>
      {!isBulk && <p style={{fontSize:12,color:"var(--color-text-secondary)",margin:"0 0 14px"}}>📄 {activity?.title}</p>}

      {/* BULK FILTERS */}
      {isBulk&&<div style={{background:"var(--color-background-secondary)",borderRadius:8,padding:14,marginBottom:14}}>
        <div style={{fontSize:13,fontWeight:600,marginBottom:8}}>{t.filterExport}</div>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>
          <div><label style={lbl}>{t.college}</label><select value={fCollege} onChange={e=>setFCollege(e.target.value)} style={inp}><option value="All">{t.all}</option>{colleges.map(c=><option key={c}>{c}</option>)}</select></div>
          <div><label style={lbl}>{t.status}</label><select value={fStatus} onChange={e=>setFStatus(e.target.value)} style={inp}><option value="All">{t.all}</option>{STATUSES.map(s=><option key={s}>{s}</option>)}</select></div>
          <div><label style={lbl}>{t.dateFrom}</label><input type="date" value={fDateFrom} onChange={e=>setFDateFrom(e.target.value)} style={inp}/></div>
          <div><label style={lbl}>{t.dateTo}</label><input type="date" value={fDateTo} onChange={e=>setFDateTo(e.target.value)} style={inp}/></div>
        </div>
        <div style={{marginTop:8,fontSize:12,fontWeight:600,color:"#1a5276"}}>{filtered.length} {t.results} {t.exportFiltered}</div>
      </div>}

      {/* Include options */}
      <div style={{fontSize:13,fontWeight:600,marginBottom:8}}>{t.includeOptions}</div>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:6,marginBottom:12}}>
        {optItems.map(o=>(
          <label key={o.k} style={{display:"flex",alignItems:"center",gap:8,fontSize:12,cursor:"pointer",padding:"6px 10px",borderRadius:8,background:opts[o.k]?"#eaf2f8":"var(--color-background-secondary)",border:`1px solid ${opts[o.k]?"#2980b9":"var(--color-border-tertiary)"}`,transition:"all 0.15s"}}>
            <input type="checkbox" checked={opts[o.k]} onChange={()=>toggle(o.k)} style={{accentColor:"#1a5276"}}/>
            {o.l}
          </label>
        ))}
      </div>
      <div style={{display:"flex",gap:8,marginBottom:16}}>
        <button onClick={allOn} style={{...btnC,fontSize:11,padding:"4px 12px"}}>{t.selectAll}</button>
        <button onClick={allOff} style={{...btnC,fontSize:11,padding:"4px 12px"}}>{t.deselectAll}</button>
      </div>

      {/* Format buttons */}
      <div style={{fontSize:13,fontWeight:600,marginBottom:10}}>{t.chooseFormat}:</div>
      <div style={{display:"flex",gap:10,flexWrap:"wrap"}}>
        <button onClick={()=>doExport("excel")} disabled={isBulk&&filtered.length===0} style={{...btnP,display:"flex",alignItems:"center",gap:6,background:"#217346",opacity:(isBulk&&filtered.length===0)?0.4:1}}>
          <span style={{fontSize:16}}>📊</span>{t.exportAsExcel}
        </button>
        <button onClick={()=>doExport("print")} disabled={isBulk&&filtered.length===0} style={{...btnP,display:"flex",alignItems:"center",gap:6,background:"#c0392b",opacity:(isBulk&&filtered.length===0)?0.4:1}}>
          <span style={{fontSize:16}}>🖨️</span>{t.exportAsPDF}
        </button>
      </div>
    </div>
  );
}

// ─── STICKY NOTES PANEL ───────────────────────────────────────────────
function StickyNotesPanel({t,notes,saveNotes,isRtl}){
  const [open,setOpen]=useState(false);
  const [text,setText]=useState("");
  const add=()=>{if(!text.trim())return;saveNotes([{id:gid(),text:text.trim(),date:new Date().toISOString().split("T")[0],time:new Date().toLocaleTimeString([],{hour:"2-digit",minute:"2-digit"})},...notes]);setText("");};
  const del=id=>saveNotes(notes.filter(n=>n.id!==id));

  return(
    <>
      {/* Floating button */}
      <button onClick={()=>setOpen(!open)} style={{position:"fixed",bottom:20,[isRtl?"left":"right"]:20,width:48,height:48,borderRadius:"50%",background:"#c9a84c",color:"#fff",border:"none",cursor:"pointer",boxShadow:"0 4px 16px rgba(0,0,0,0.2)",display:"flex",alignItems:"center",justifyContent:"center",zIndex:90,fontSize:20,transition:"transform 0.2s",transform:open?"rotate(45deg)":"none"}}>
        {open?"✕":"📌"}
      </button>

      {/* Panel */}
      {open&&<div style={{position:"fixed",bottom:78,[isRtl?"left":"right"]:20,width:320,maxHeight:420,background:"#ffffff",borderRadius:14,border:"0.5px solid var(--color-border-tertiary)",boxShadow:"0 8px 32px rgba(0,0,0,0.12)",zIndex:90,display:"flex",flexDirection:"column",overflow:"hidden"}}>
        <div style={{padding:"14px 16px 10px",borderBottom:"0.5px solid var(--color-border-tertiary)",display:"flex",alignItems:"center",justifyContent:"space-between"}}>
          <span style={{fontSize:14,fontWeight:700,color:"#c9a84c"}}>📌 {t.stickyNotes}</span>
          <span style={{fontSize:11,color:"var(--color-text-tertiary)"}}>{notes.length}</span>
        </div>

        {/* Add note input */}
        <div style={{padding:"10px 16px",borderBottom:"0.5px solid var(--color-border-tertiary)",display:"flex",gap:8}}>
          <textarea value={text} onChange={e=>setText(e.target.value)} onKeyDown={e=>{if(e.key==="Enter"&&!e.shiftKey){e.preventDefault();add();}}} placeholder={t.noteText} rows={2} style={{...inp,flex:1,resize:"none",fontSize:12,padding:"6px 10px"}}/>
          <button onClick={add} disabled={!text.trim()} style={{...btnP,padding:"8px 12px",fontSize:11,alignSelf:"flex-end",opacity:text.trim()?1:0.4}}>{t.addNote}</button>
        </div>

        {/* Notes list */}
        <div style={{flex:1,overflowY:"auto",padding:"8px 16px"}}>
          {notes.length===0&&<div style={{textAlign:"center",padding:20,color:"var(--color-text-tertiary)",fontSize:12}}>{t.noNotes}</div>}
          {notes.map(n=>(
            <div key={n.id} style={{background:"#fef9e7",borderRadius:8,padding:"10px 12px",marginBottom:8,position:"relative",borderLeft:"3px solid #c9a84c"}}>
              <div style={{fontSize:12,color:"#2c3e50",lineHeight:1.6,whiteSpace:"pre-wrap"}}>{n.text}</div>
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginTop:6}}>
                <span style={{fontSize:10,color:"#b7950b"}}>{n.date} · {n.time}</span>
                <button onClick={()=>del(n.id)} style={{background:"none",border:"none",cursor:"pointer",padding:2,fontSize:12,color:"#e74c3c"}} title={t.deleteNote}>✕</button>
              </div>
            </div>
          ))}
        </div>
      </div>}
    </>
  );
}

// ─── HELP PANEL ──────────────────────────────────────────────────────
function HelpPanel({page,isRtl,lang}){
  const [open,setOpen]=useState(false);
  const ar=lang==="ar";

  const help={
    dashboard:{
      title:ar?"لوحة التحكم":"Dashboard",
      sections:[
        {h:ar?"المقاييس القابلة للنقر":"Clickable metrics",t:ar?"اضغط على أي بطاقة مقياس (الأنشطة، المكتملة، المتأخرة...) لعرض القائمة التفصيلية. اضغط مرة أخرى للإغفال.":"Click any stat card (Total, Completed, Overdue, etc.) to see the detailed list. Click again to close."},
        {h:ar?"فلتر الكليات":"College filter",t:ar?"استخدم القائمة المنسدلة في الأعلى لعرض بيانات كلية واحدة. جميع المخططات والمقاييس تتحدث معاً.":"Use the dropdown at top-right to view one college's data. All charts and metrics update together."},
        {h:ar?"الأنشطة المتأخرة":"Overdue tracking",t:ar?"البطاقة الحمراء تظهر الأنشطة التي تجاوزت موعدها ولم تكتمل. اضغط عليها لمعرفة أيها بالضبط.":"The red card shows activities past their date but not completed. Click it to see exactly which ones."},
        {h:ar?"الحضور":"Attendance",t:ar?"أرقام الحضور تتزامن تلقائياً من تقارير الأحداث. قدّم تقرير حدث لتحديث العدد.":"Attendance numbers sync from event reports. File an event report to update the count."},
      ]
    },
    plans:{
      title:ar?"الخطط السنوية":"Yearly Plans",
      sections:[
        {h:ar?"سير العمل":"Workflow",t:ar?"1. أنشئ خطة لكل كلية\n2. أضف الأنشطة يدوياً أو استخرجها بالذكاء الاصطناعي\n3. وافق على الخطة\n4. تابع التنفيذ":"1. Create a plan for each college\n2. Add activities manually or AI-extract from uploaded documents\n3. Approve the plan\n4. Track execution"},
        {h:ar?"الاستيراد بالذكاء الاصطناعي":"AI bulk import",t:ar?"ارفع صورة أو PDF لخطة سنوية → النظام يستخرج جميع الأنشطة دفعة واحدة. راجعها قبل الاستيراد.":"Upload a photo or PDF of a yearly plan → the system extracts all activities at once. Review before importing."},
        {h:ar?"السنة الأكاديمية":"Academic year",t:ar?"سبتمبر-ديسمبر = 2025، يناير-أغسطس = 2026. النظام يحدد السنة تلقائياً.":"Sep-Dec = 2025, Jan-Aug = 2026. The system auto-assigns the correct year."},
        {h:ar?"التصدير":"Export",t:ar?"كل نشاط له زر تصدير → إكسل أو PDF مع خيارات تخصيص.":"Each activity has an export button → Excel or PDF with customization options."},
      ]
    },
    activities:{
      title:ar?"جميع الأنشطة":"All Activities",
      sections:[
        {h:ar?"الفلاتر":"Filters",t:ar?"فلتر بالمصدر، الحالة، الكلية، أو البحث النصي. الفلاتر تؤثر على التصدير أيضاً.":"Filter by source, status, college, or text search. Filters apply to exports too."},
        {h:ar?"التصدير المجمع":"Bulk export",t:ar?"زر 'تصدير الكل' يصدر فقط الأنشطة المعروضة (بعد الفلترة). اختر الأقسام التي تريدها.":"'Export All' exports only displayed activities (after filtering). Choose which sections to include."},
      ]
    },
    initiatives:{
      title:ar?"المبادرات":"Initiatives",
      sections:[
        {h:ar?"أنواع المبادرات":"Initiative types",t:ar?"مبادرة القسم، مبادرة الجامعة، توجيه الوزارة، أو طلب كلية. كلها تُتبع هنا.":"Department Initiative, University Initiative, Ministry Directive, or College Request. All tracked here."},
        {h:ar?"الربط بالكلية":"College linking",t:ar?"كل مبادرة يمكن ربطها بكلية محددة أو تكون على مستوى الجامعة.":"Each initiative can be linked to a specific college or be university-wide."},
        {h:ar?"تقرير الحدث":"Event report",t:ar?"عند تقديم تقرير حدث لمبادرة، يتحدث الحضور والحالة تلقائياً.":"When you file an event report for an initiative, attendance and status auto-sync."},
      ]
    },
    directives:{
      title:ar?"التوجيهات والأوامر":"Directives & Orders",
      sections:[
        {h:ar?"سير العمل المثالي":"Practical workflow",t:ar?"1. أنشئ توجيه جديد (مثلاً: أمر وزاري)\n2. حدد الكليات المستهدفة (كلها أو محددة)\n3. اضغط ⚡ 'إنشاء مبادرة لكل كلية' → يُنشئ مبادرة لكل كلية تلقائياً\n4. الكليات تنفذ وتقدم تقارير → الامتثال يتحدث تلقائياً\n5. تابع نسبة الامتثال وتواصل مع المتأخرين":"1. Create a new directive (e.g., ministry order)\n2. Select target colleges (all or specific)\n3. Click ⚡ 'Create initiative for each college' → auto-creates initiatives\n4. Colleges execute and file reports → compliance auto-updates\n5. Track compliance % and follow up with laggards"},
        {h:ar?"الامتثال التلقائي":"Auto-compliance",t:ar?"النظام يبحث تلقائياً عن أنشطة مطابقة في كل كلية. إذا وجد نشاطاً بعنوان مشابه، يربطه تلقائياً.":"The system auto-scans for matching activities in each college. If it finds one with a similar title, it links automatically."},
        {h:ar?"الربط اليدوي":"Manual link",t:ar?"إذا لم يجد النظام تطابقاً، اضغط 'ربط نشاط' واختر النشاط المناسب من قائمة الكلية.":"If auto-match fails, click 'Link activity' and select the right one from the college's list."},
        {h:ar?"تحديد كمكتمل":"Mark compliant",t:ar?"يمكنك تحديد كلية كمكتملة يدوياً إذا أكملت المهمة خارج النظام.":"You can manually mark a college as compliant if they completed the task outside the system."},
      ]
    },
    trainers:{
      title:ar?"المدربين":"Trainer Profiles",
      sections:[
        {h:ar?"ربط الأنشطة":"Link activities",t:ar?"عند إضافة مدرب، اختر الأنشطة التي شارك فيها من القائمة المنسدلة. تظهر كوسوم زرقاء.":"When adding a trainer, select activities they participated in from the dropdown. They appear as blue tags."},
        {h:ar?"الاستخراج":"AI extraction",t:ar?"ارفع CV أو نموذج ترشيح → النظام يستخرج الاسم والتخصص والمعلومات تلقائياً.":"Upload a CV or nomination form → the system extracts name, specialization, and info automatically."},
        {h:ar?"التصدير":"Export",t:ar?"إكسل أو PDF. يتضمن جميع المعلومات مع الفلترة بالكلية.":"Excel or PDF. Includes all info filtered by college."},
      ]
    },
    staff:{
      title:ar?"كادر الكليات":"College Staff",
      sections:[
        {h:ar?"مصادر البيانات":"Data sources",t:ar?"الكادر يظهر من مصدرين: 1) تلقائياً من المحاضرين في الأنشطة 2) استيراد قائمة الجامعة":"Staff appears from two sources: 1) Auto-detected from activity presenters 2) Imported university roster"},
        {h:ar?"الاستيراد":"Import",t:ar?"ارفع ملف Excel أو CSV أو نص عادي. النظام يكتشف الأعمدة تلقائياً ويمنع التكرار بمقارنة ذكية للأسماء.":"Upload Excel, CSV, or plain text. System auto-detects columns and prevents duplicates with fuzzy name matching."},
        {h:ar?"كتب الشكر":"Appreciation letters",t:ar?"يمكن إضافتها يدوياً أو يكتشفها النظام تلقائياً من المراسلات التي تحتوي على كلمة 'شكر'.":"Can be added manually or auto-detected from correspondence containing 'شكر' or 'appreciation'."},
        {h:ar?"فلتر 'لا نشاط'":"'No activity' filter",t:ar?"يظهر الكادر المستورد الذين لم يشاركوا في أي نشاط — مفيد لمتابعة الالتزام.":"Shows imported staff with zero participation — useful for tracking compliance."},
      ]
    },
    reports:{
      title:ar?"تقارير الأحداث":"Event Reports",
      sections:[
        {h:ar?"سير العمل":"Workflow",t:ar?"1. اضغط 'تقرير جديد'\n2. اختر النشاط من القائمة (يشمل الخطط والمبادرات)\n3. البيانات تُملأ تلقائياً\n4. أضف عدد المشاركين وأسماءهم والملخص\n5. حفظ → النشاط يتحدث تلقائياً كمكتمل":"1. Click 'New Report'\n2. Select activity from dropdown (includes plans + initiatives)\n3. Data auto-fills\n4. Add participant count, names, and summary\n5. Save → activity auto-updates to Completed"},
        {h:ar?"المزامنة":"Sync",t:ar?"تقديم التقرير يحدث تلقائياً: عدد الحضور، أسماء المشاركين، والحالة تصبح 'مكتمل'.":"Filing a report auto-syncs: attendee count, participant names, and status becomes 'Completed'."},
        {h:ar?"التعديل":"Editing",t:ar?"التقارير قابلة للتعديل بعد الحفظ. التغييرات تتزامن مع النشاط المرتبط.":"Reports are editable after saving. Changes re-sync to the linked activity."},
        {h:ar?"التقارير لا تختفي":"Reports persist",t:ar?"التقارير المكتملة تبقى دائماً في القائمة. لا تختفي بعد الإكمال.":"Completed reports stay in the list forever. They don't disappear after completion."},
      ]
    },
    correspondence:{
      title:ar?"المراسلات":"Correspondence",
      sections:[
        {h:ar?"البحث":"Search",t:ar?"شريط البحث يبحث في كل الحقول: الموضوع، الرقم، المرسل، المستلم، الملاحظات، التاريخ، الكلية، الحقول المخصصة.":"Search bar scans all fields: subject, ref number, sender, recipient, notes, date, college, custom fields."},
        {h:ar?"الملفات":"File attachments",t:ar?"الصور تُضغط تلقائياً عند الرفع (~90% أصغر). يمكنك حذف الملف من البطاقة مباشرة.":"Images auto-compress on upload (~90% smaller). You can delete the file directly from the card."},
        {h:ar?"الاستخراج":"AI extraction",t:ar?"ارفع صورة أو PDF لكتاب رسمي → النظام يستخرج المعلومات تلقائياً.":"Upload a photo or PDF of an official letter → system extracts info automatically."},
      ]
    },
    analytics:{
      title:ar?"التحليلات":"Analytics",
      sections:[
        {h:ar?"التصدير المخصص":"Custom export",t:ar?"اضغط 📤 تصدير → اختر الأقسام التي تريد تضمينها → إكسل أو PDF.":"Click 📤 Export → check/uncheck sections you want → Excel or PDF."},
        {h:ar?"مقارنة الكليات":"College comparison",t:ar?"جدول الأداء يقارن: الأنشطة، الخطط مقابل المبادرات، نسبة الإكمال، المتأخر، بدون تقرير، الحضور.":"Performance table compares: activities, plans vs initiatives, completion rate, overdue, no-report, attendance."},
        {h:ar?"بدون تقرير":"'No report' detection",t:ar?"يكشف الأنشطة المحددة 'مكتملة' بدون تقديم تقرير حدث — مشكلة سلامة بيانات.":"Detects activities marked 'Completed' without filing an event report — a data integrity issue."},
        {h:ar?"التغطية الإعلامية":"Media coverage",t:ar?"يتتبع حالة الإعلام لكل نشاط: منشور، تم التواصل، مرفوض، لم يتم التواصل. مع التصدير.":"Tracks media status per activity: published, contacted, declined, not contacted. With export."},
      ]
    },
    backup:{
      title:ar?"النسخ الاحتياطي":"Backup & Export",
      sections:[
        {h:ar?"الأرشيف":"Archive system",t:ar?"انقل السنوات المكتملة إلى تخزين أرشيفي لتحرير المساحة. كل سنة تحصل على 5MB خاصة.":"Move completed years to archive storage to free space. Each year gets its own 5MB."},
        {h:ar?"التصدير الانتقائي":"Selective export",t:ar?"اختر الفئات والفلاتر بالضبط ثم صدّر.":"Choose exactly which categories and filters, then export."},
        {h:ar?"إدارة الملفات":"File management",t:ar?"حدد الملفات للتنزيل أو الحذف بشكل منفصل. نزّل أولاً، تحقق، ثم احذف.":"Select files to download or delete separately. Download first, verify, then clear."},
      ]
    },
    settings:{
      title:ar?"الإعدادات":"Settings",
      sections:[
        {h:ar?"العلامة التجارية":"Branding",t:ar?"غيّر اسم المنصة، العنوان الفرعي، والشعار. يظهر في الشريط الجانبي وجميع التصديرات.":"Change platform name, subtitle, and logo. Appears in sidebar and all exports."},
        {h:ar?"إدارة الكليات":"College management",t:ar?"أضف، عدّل، أو احذف كليات. التعديل يُحدّث الاسم في جميع البيانات المرتبطة تلقائياً.":"Add, edit, or delete colleges. Editing auto-renames across all linked data."},
        {h:ar?"تغيير كلمة المرور":"Change password",t:ar?"أدخل كلمة المرور الحالية ثم الجديدة. مشفّرة SHA-256.":"Enter current then new password. SHA-256 hashed."},
      ]
    },
  };

  const h=help[page]||{title:ar?"المساعدة":"Help",sections:[{h:ar?"مرحباً":"Welcome",t:ar?"اختر صفحة من الشريط الجانبي للحصول على مساعدة سياقية.":"Select a page from the sidebar to get contextual help."}]};

  return(
    <>
      <button onClick={()=>setOpen(!open)} style={{position:"fixed",bottom:20,[isRtl?"left":"right"]:76,width:48,height:48,borderRadius:"50%",background:"#2980b9",color:"#fff",border:"none",cursor:"pointer",boxShadow:"0 4px 16px rgba(0,0,0,0.2)",display:"flex",alignItems:"center",justifyContent:"center",zIndex:90,fontSize:20,transition:"transform 0.2s",transform:open?"rotate(45deg)":"none"}}>
        {open?"✕":"❓"}
      </button>
      {open&&<div dir={ar?"rtl":"ltr"} style={{position:"fixed",bottom:78,[isRtl?"left":"right"]:76,width:340,maxHeight:480,background:"#ffffff",borderRadius:14,border:"0.5px solid var(--color-border-tertiary)",boxShadow:"0 8px 32px rgba(0,0,0,0.12)",zIndex:90,display:"flex",flexDirection:"column",overflow:"hidden"}}>
        <div style={{padding:"14px 16px 10px",borderBottom:"0.5px solid var(--color-border-tertiary)",background:"#eaf2f8"}}>
          <span style={{fontSize:14,fontWeight:700,color:"#1a5276"}}>❓ {h.title}</span>
        </div>
        <div style={{flex:1,overflowY:"auto",padding:"10px 16px"}}>
          {h.sections.map((s,i)=>(
            <div key={i} style={{marginBottom:12}}>
              <div style={{fontSize:12,fontWeight:700,color:"#1a5276",marginBottom:3}}>{s.h}</div>
              <div style={{fontSize:11,color:"#2c3e50",lineHeight:1.8,whiteSpace:"pre-line"}}>{s.t}</div>
            </div>
          ))}
          <div style={{marginTop:8,padding:"8px 10px",background:"#fef9e7",borderRadius:6,fontSize:10,color:"#7d6608",lineHeight:1.7}}>
            💡 {ar?"نصيحة: استخدم زر المزامنة 🔄 في الشريط العلوي إذا كنت تعمل على أكثر من جهاز.":"Tip: Use the sync 🔄 button in the top bar if working across multiple devices."}
          </div>
        </div>
      </div>}
    </>
  );
}

// ─── GENERIC COMPONENTS ───────────────────────────────────────────────
function FrmCard({t,title,onCancel,onSave,fields}){
  const [d,sD]=useState(Object.fromEntries(fields.map(f=>[f.k,f.v||""])));
  const s=(k,v)=>sD({...d,[k]:v});
  return(
    <div style={{...crd,border:"1.5px solid #c9a84c",marginBottom:16}}>
      <h3 style={{fontSize:14,fontWeight:600,margin:"0 0 12px"}}>{title}</h3>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>
        {fields.map(f=><div key={f.k}><label style={lbl}>{f.l}</label>{f.type==="fixed"?<input value={d[f.k]} disabled style={{...inp,background:"var(--color-background-secondary)"}}/>:f.type==="select"?<select value={d[f.k]} onChange={e=>s(f.k,e.target.value)} style={inp}><option value="">--</option>{f.opts.map(o=><option key={o}>{o}</option>)}</select>:<input value={d[f.k]} onChange={e=>s(f.k,e.target.value)} style={inp}/>}</div>)}
      </div>
      <div style={{display:"flex",gap:8,marginTop:12,justifyContent:"flex-end"}}><button onClick={onCancel} style={btnC}>{t.cancel}</button><button onClick={()=>onSave(d)} style={btnP}>{t.save}</button></div>
    </div>
  );
}

function Flt({l,v,set,opts}){return <div style={{display:"flex",alignItems:"center",gap:4}}><span style={{fontSize:10,color:"var(--color-text-tertiary)"}}>{l}:</span><select value={v} onChange={e=>set(e.target.value)} style={{padding:"4px 8px",border:"1px solid var(--color-border-tertiary)",borderRadius:6,fontSize:11,background:"#ffffff",color:"var(--color-text-primary)"}}>{opts.map(o=><option key={o}>{o}</option>)}</select></div>;}
