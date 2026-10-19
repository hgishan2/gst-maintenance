// ===== EDIT THIS FILE (only SCRIPT_URL and KEY are required) =====
const CONFIG = {
  SCRIPT_URL: "https://script.google.com/macros/s/AKfycbxUX47oeLozvleYidP0xmSUsz5YDonxJOqSo4oxvwcvsx9AMcOd9QLfaj5ma1_xOZk3-Q/exec",              // paste your Google Apps Script Web App URL here
  KEY: "GST12345",   // must match SECRET in Code.gs

  SUPERVISOR: "Ishan Madhushanka",   // pre-filled supervisor (technicians can still change it)
  BUILDING: "GST1",            // fixed building name (technicians cannot change it)
  TECHNICIANS: ["Nagendra Vishwakarma", "Ajeesh Mohankumar", "Kaushar Ali", "Mehedi Hassan", "Ram Jatan", "Mohammad Jahangeer", "Ali Hassan", "Imtiyaz Kamruddin", "Madhab Charan"],   // dropdown names (the "Lists" tab in the Google Sheet overrides this)

  // Details printed on each overtime sheet (name = as shown in the app's dropdown)
  STAFF: [
    {name:"Mehedi Hassan",        code:"3765",   designation:"Mason/Painter",      full:"MEHEDI HASAN LIYAKAT ALI"},
    {name:"Kaushar Ali",          code:"711312", designation:"Plumber",            full:"MOHAMMAD KAUSHAR ALI"},
    {name:"Ajeesh Mohankumar",    code:"712371", designation:"Painter",            full:"AJEESH MOHANKUMAR GEETHA"},
    {name:"Nagendra Vishwakarma", code:"712892", designation:"General Technician", full:"NAGENDRA VISHWAKARMA"},
    {name:"Ram Jatan",            code:"713052", designation:"Carpenter",          full:"RAM JATAN MAHADEV"},
    {name:"Mohammad Jahangeer",   code:"713150", designation:"AC Technician",      full:"MUHAMMAD JAHANGEER"},
    {name:"Ali Hassan",           code:"713502", designation:"Electrician",        full:"ALI HASSAN"},
    {name:"Imtiyaz Kamruddin",    code:"713520", designation:"AC Technician",      full:"IMTIYAZ KAMRUDDIN"},
    {name:"Madhab Charan",        code:"713542",       designation:"Plumber",            full:"MADHAB CHARAN"}
  ],

  // Utility meters (after first use, the "Meters" tab in the Google Sheet is the master list)
  METERS: [
    {type:"Electricity", area:"Common",     meter:"21172699",  account:"2888542999"},
    {type:"Electricity", area:"101",        meter:"21135023",  account:"4254023806"},
    {type:"Electricity", area:"205",        meter:"21135492",  account:"20000282308"},
    {type:"Water",       area:"Common",     meter:"10W708361", account:"2888542999"},
    {type:"Water",       area:"101",        meter:"12A009001", account:"4254023806"},
    {type:"Water",       area:"205",        meter:"12A009016", account:"20000282308"},
    {type:"Water",       area:"Fire Water", meter:"110194605", account:"7004047472"}
  ],

  // Same lists as your Excel dropdowns
  CATEGORIES: ["AC / HVAC","Plumbing","Electrical","Civil / Painting","Fire & Life Safety","Carpentry","General Maintenance","PPM","Other"],
  PRIORITIES: ["Emergency","High","Medium","Low"],
  STATUSES: ["Completed","Pending","In Progress","On Hold"],

  // Extra auto-corrections for the text boxes: "what they type": "what it becomes" (lowercase)
  CORRECTIONS: {}
};
