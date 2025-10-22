# 🧱 The Wall — Interactive Data Installation (Google Apps Script)

> **Summary:**  
> The Wall is a modular, data-driven visualization environment built entirely with **Google Apps Script (V8)**, **HTML/CSS/JS**, and **Google Sheets**.  
> It transforms form-submitted social data into real-time visual dashboards: **Check-In**, **Wall (Demographics)**, **Matchmaker**, **Map**, **Intro**, and **Display Controller**.  
> This README describes every component, expected backend interfaces, and architectural rules so automated tools or AI refactors can fix bugs **without altering the functional design**.

---

## 🧭 Core Concept
The project connects a live Google Sheet to a rotating display of analytics pages.  
Each page (`intro.html`, `wall.html`, `mm.html`, `map.html`, etc.) is requested through  
`https://script.google.com/.../exec?page=NAME`.  
A central `Display.html` cycles through these stages automatically.

---

## 🧩 Architecture Overview

Google Form → Google Sheet
↓
Apps Script Backend (.gs)
• Code.gs → router + global utility
• Analytics.gs → compatibility + data transforms
• onOpen.gs → custom menu & triggers
↓
HTML Frontend (.html)
• CheckInInterface.html → guest check-in & photo upload
• Display.html → master controller
• intro.html → intro animation
• wall.html → demographic visualization
• mm.html → compatibility/matchmaker
• map.html → zip-based network map
↓
appsscript.json → project manifest


---

## 🧱 File-by-File Purpose

### **appsscript.json**
- Runtime: `V8`
- Exception logging: `STACKDRIVER`
- WebApp access: `ANYONE`, `executeAs: USER_DEPLOYING`
- Enables the **Maps** advanced service (v2).

### **Code.gs**
- Defines the `doGet(e)` router:
  ```js
  function doGet(e) {
    const page = e.parameter.page || 'intro';
    return HtmlService.createHtmlOutputFromFile(page)
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
