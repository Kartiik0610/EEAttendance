# Echaar Attendance — Geofenced Workforce Attendance Management System

A complete, responsive, mobile-first Employee Attendance Management Web Application built with **HTML5**, **Tailwind CSS**, **Vanilla JavaScript**, and a **Google Apps Script (GAS) Web App** backend connected to a single Google Spreadsheet.

---

## 🌟 Key Features

1. **Dual-Tab Google Sheets Architecture**:
   - **`Attendance` Tab**: Records `Date (YYYY-MM-DD)`, `Employee_ID`, `Employee_Name`, `Punch_In_Time`, `Punch_Out_Time`, `Status` (`Present` / `Absent`), `Absence_Reason`, `Device_Fingerprint`, and `Location_Coords`.
   - **`Users` Tab**: Manages user directory with `Employee_ID`, `Employee_Name`, `Phone_Number` (User ID), `Password` (`pass@123`), and `Role` (`Employee` or `Admin`).
   - Automated 1-click sheet builder (`initialSetup()`) that formats headers and pre-populates sample data.
   - **Tamper-Proof Audit Trail**: Punch timestamps are immutable logs that cannot be altered or deleted.

2. **Main Landing Page (Quick Punch Interface)**:
   - Modern, mobile-first responsive card layout with live ticking digital clock (HH:MM:SS AM/PM).
   - Prominent, auto-capitalizing **Employee ID** input (`e.g., EEI001`).
   - Clean action buttons:
     - **Punch In** (Green): Validates geofence and records timestamp, coordinates, and device fingerprint.
     - **Punch Out** (Amber/Red): Finds today's row and updates departure timestamp.
     - **Reason for Absent** (Slate Grey): Opens a modal with presets and remarks for absence logging.

3. **Strict Geofencing & Security**:
   - **GPS Geofencing**: Computes distance using the **Haversine formula** from office coordinates (`19.193740, 72.957355`, allowed radius `80m`). Blocks Punch In/Out if outside allowed radius.
   - **One Punch In & One Punch Out Per Day**: Enforces strict single-entry logging per employee per day on both client and backend.
   - **One Device Per Day Lock**: Generates a composite hardware/browser fingerprint (LocalStorage UUID + Screen resolution + Timezone + UserAgent) and prevents one physical device from clocking in for multiple employees on the same day.

4. **Role-Based Authentication & Views**:
   - **Employee Portal**:
     - Interactive **Monthly Calendar View** with month/year navigation.
     - Visual Color Coding:
       - 🟢 **Green**: `Present` (shows punch in/out times on click).
       - 🔴 **Red**: `Absent` (shows logged reason on click).
       - 🔵 **Blue**: **Office Week Off — Every Friday** (all Fridays automatically highlighted in Blue).
       - ⚪ **Grey**: Future dates or unlogged working days.
     - Monthly KPI counters: Total Present Days, Total Absent Days, Friday Offs, and Monthly Attendance Rate (%).
   - **Admin / MD Dashboard**:
     - Real-time Top KPI Cards: Total Employees, Present Today, Absent Today, Not Yet Punched In.
     - **Daily Roster Table**: Filterable by status (`All`, `Present`, `On-Time`, `Late`, `Absent`, `Not Punched In`) and searchable by name, ID, or phone. Includes work duration calculation and punctuality score against official shift (e.g., 09:30 AM).
     - **Monthly Organization Calendar**: Organization-wide monthly calendar displaying total present counts with employee names, absent counts with absence reasons, and Friday week-offs.
     - **Password Management**: Change password modal allowing users to change their own password, and admins to reset staff passwords.
     - **One-click Export to CSV** for payroll/HR reporting.

---

## 📁 Project Structure

```
Attendance/
├── Code.gs        # Google Apps Script backend (Web App API endpoints & Sheet setup)
├── index.html     # Responsive Single-Page Application (Tailwind CSS, Icons, JS)
├── logo.png       # Echaar company brand logo
├── server.js      # Local Node.js static server with strict cache-control headers
└── README.md      # Setup, deployment, and configuration guide
```

---

## 🚀 Setup & Deployment Guide

### Google Apps Script Web App Deployment
* **Backend Script**: Paste [`Code.gs`](file:///d:/FOLDER/B.S/Attendance/Code.gs) into Google Apps Script connected to your Google Sheet.
* **Run `initialSetup()`**: Automatically creates and formats the `Attendance` and `Users` tabs.
* **Deploy as Web App**:
  - Execute as: **Me**
  - Who has access: **Anyone**

### Office Geofence Coordinates
* **Latitude**: `19.193740`
* **Longitude**: `72.957355`
* **Allowed Radius**: `80 meters`

---

## 👥 Default Credentials

| Role | Employee ID | Phone Number (User ID) | Default Password |
| :--- | :--- | :--- | :--- |
| **Admin / MD** | `EEI001` | `8888888888` | `pass@123` |
| **Employee** | `EEI002` | `9876543211` | `pass@123` |
