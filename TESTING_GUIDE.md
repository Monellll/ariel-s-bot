# Ariel's Bot — Self-Testing Guide & Sample Data Generator

> **Purpose:** This guide walks you through generating sample test data and testing every feature of the Telegram Payment Reminder Bot yourself on your own device.

---

## 🚀 Quick Start (Testing Environment)

The bot is currently in **Testing Mode**:
* **Check Interval:** Runs automatically **every 3 minutes**.
* **Target Offset:** `REMINDER_DAYS_BEFORE = 0` (sends reminders for students whose payment is due **today**).
* **Timezone:** `Asia/Yangon` (UTC+6:30).

---

## 🛠️ Step 1: Generate Your Test Excel File

We've provided a helper script [`create_test_xlsx.py`](file:///C:/Users/Mon%20El/OneDrive/Desktop/Ariel's%20Bot/telegram-payment-reminder/create_test_xlsx.py) that generates a test spreadsheet with dates matching **today** in `Asia/Yangon`.

Open PowerShell and run:

```powershell
cd "C:\Users\Mon El\OneDrive\Desktop\Ariel's Bot\telegram-payment-reminder"
.\venv\Scripts\Activate

# Generate with your own Telegram username in Row 1:
python create_test_xlsx.py @YourTelegramUsername
```

### What this file creates (`test_upload.xlsx`):

| Row | Username | Student Name | Due Date | Start Date | Scenario Tested |
|---|---|---|---|---|---|
| **1** | `@YourTelegramUsername` | Myat (Self Test) | **Today** | `2026-07-14` | Immediate delivery to your phone |
| **2** | `@arealtest` | Bao (Test) | **Today** | `2026-08-01` | Dynamic custom start date |
| **3** | `@thiri_test` | Thiri (Upcoming) | **In 5 days** | `2026-07-14` | Upcoming reminders queue |
| **4** | `@kyi_test` | Kyi (Future) | **In 12 days** | _(default)_ | Future reminders queue |

The file is saved at:
* [`C:\Users\Mon El\OneDrive\Desktop\Ariel's Bot\test_upload.xlsx`](file:///C:/Users/Mon%20El/OneDrive/Desktop/Ariel's%20Bot/test_upload.xlsx)

---

## 🤖 Step 2: Start the Bot

In PowerShell (with virtual environment active):

```powershell
python bot.py
```

Leave this terminal window running while testing.

---

## 🧪 Step 3: Self-Testing Checklist (7 Live Tests)

### Test 1: Student Auto-Registration
1. Open your bot in Telegram from your personal account.
2. Send:
   ```text
   /start
   ```
3. **Expected Result:**
   * If you are the Admin/Teacher: Displays the new **Interactive Control Panel (Menu UI)**.
   * If testing from a student account: Displays:
     > *"✅ You are now registered! You will receive automatic payment reminders when your teacher schedules them."*
   * Your `@username` and numeric `chat_id` are saved into `data/chat_id_registry.json`.

---

### Test 2: Upload the Excel File
1. In the bot chat, attach and send `test_upload.xlsx` as a document.
2. **Expected Result:**
   * The bot validates the document.
   * Reports total records (4), valid records (4), and auto-matched Chat IDs.
   * Immediately runs a check for students due today and displays:
     > *"📨 Immediately sent 1 reminder(s)"*

---

### Test 3: Verify the Dynamic Bilingual Reminder Message
Check the private message delivered to your Telegram account.

**Expected Message Format:**
```text
Hi! 

Just a friendly reminder that since we started our class on July 14th , the fee for this month was due on October 17th.

No worries at all, just a gentle reminder.

Take your time! 🤍

Hello, ချယ်ချယ်တို့ အတန်းလေးက July 14 ရက်နေ့က စတာမို့ October 17 ရက်နေ့ကျ သင်တန်းကြေးလေး ပြန်သွင်းပေးရပါမယ်ရှင့်

မလောပါဘူးရှင် remind လုပ်တာပါနော်🤍
```

* **Dynamic Dates Verified:** Notice that `October 17th` in English and `October 17 ရက်နေ့` in Myanmar match today's date automatically.

---

### Test 4: Navigate the Interactive Menu UI
Send `/menu` (or `/start`):

* Tap **`[📊 Roster & Status]`**:
  * Shows loaded student count (4), Chat IDs linked vs. missing, and next upcoming deadlines.
* Tap **`[⚙️ Scheduler]`**:
  * Shows scheduler engine status (`🟢 Active`, every 3 minutes).
  * Shows live countdown timer (e.g. `⏳ in 1m 45s`).
  * Shows **Next Check Queue** and **Upcoming Scheduled Queue**.
* Tap **`[🔄 Refresh Menu]`**:
  * Watch the countdown timer update in real-time.

---

### Test 5: Verify Duplication is Impossible
1. Tap the **`[📨 Check & Send Now]`** button (or send `/check`).
2. **Expected Result:**
   * Output reports:
     ```text
     ℹ️ Already Sent (1):
       • Myat (@YourUsername) — due 2026-10-17
     ```
   * **No duplicate message is sent to your Telegram!**
3. Wait for the 3-minute timer to expire:
   * The scheduler runs automatically in the background.
   * Checks `sent_reminders.json` on disk and skips re-sending.

---

### Test 6: Delete an Individual Reminder Schedule
1. In the menu, tap **`[🗑️ Delete Reminders]`** (or send `/delete`).
2. You will see individual delete buttons for each student:
   ```text
   [❌ 1. @YourUsername]
   [❌ 2. @arealtest]
   [❌ 3. @thiri_test]
   [❌ 4. @kyi_test]
   ```
3. Tap **`[❌ 2. @arealtest]`**.
4. Tap **`[🗑️ Yes, Delete]`** to confirm.
5. **Expected Result:**
   * Shows confirmation: `✅ Deleted reminder for @arealtest!`.
   * The student is removed from the active in-memory schedule.
   * The user's `data/{your_id}/students.xlsx` file is automatically rewritten to disk, persisting the deletion across restarts.

---

### Test 7: Reset History & Test Sending Again
1. In the menu, tap **`[📜 Sent History]`**.
   * Shows your recorded `chat_id:due_date` entry.
2. Tap **`[🗑️ Clear Sent History]`** (or send `/clearsent`).
   * History is cleared.
3. Tap **`[📨 Check & Send Now]`**:
   * The bot re-evaluates today's deadline and delivers the reminder to your phone again.

---

## 🛡️ Deduplication Verification Summary

| Protection Layer | Test Method | Result |
|---|---|---|
| **Concurrency Lock** | Trigger manual check while 3-min timer fires | Serialized cleanly by `asyncio.Lock` |
| **Atomic File Flush** | Re-check `sent_reminders.json` on disk | Flushed via `os.replace` immediately on each send |
| **Duplicate Rows in Excel** | Paste same student twice on same date | Automatically deduplicated on file load |
| **Memory Cache Fallback** | Test with simulated file read errors | Protected by in-memory `_memory_sent_cache` |

---

## ⚙️ Switching from Testing Mode to Production

When you finish self-testing and are ready for production:

1. In `.env`:
   ```ini
   REMINDER_DAYS_BEFORE=5
   ```
2. In `bot.py` (lines ~1370):
   Change APScheduler from interval to daily cron at 08:00 AM:
   ```python
   _scheduler.add_job(
       _run_daily_check,
       trigger="cron",
       hour=8,
       minute=0,
       kwargs={"bot": application.bot},
       id="daily_reminder_check",
       replace_existing=True,
   )
   ```
