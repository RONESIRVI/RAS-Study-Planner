# RAS Study Automation System

Automated study planner and reminder system for RAS preparation.

## Features
- **Adaptive Study Plan**: Generates daily topics based on progress tracker.
- **PYQ Generation**: Automatically creates Excel files for Previous Year Questions.
- **Premium Image Reports**: Visualizes daily schedule.
- **Weekend System**: Automatically switches to Revision/Reflection mode on Saturdays and Sundays.
- **Configurable Scheduling**: Set your preferred emailing times.

## Configuration (`config.json`)
You can customize the system behavior in `config.json`:

```json
{
    "sender_email": "your-email@gmail.com",
    "sender_password": "your-app-password",
    "recipient_email": "recipient@gmail.com",
    "weekend_mode": true,
    "morning_email_time": "08:00",
    "evening_email_time": "21:30"
}
```

- **`weekend_mode`**: Set to `true` to enable specialized weekend messages (Revision & Reflection). Set to `false` for regular study plans every day.
- **`morning_email_time`**: The target time for the morning study plan email (24h format).
- **`evening_email_time`**: The target time for the evening reminder email (24h format).

## Execution
The system is designed to be triggered by Task Scheduler using the provided batch files:
- `Run_Morning_Plan.bat`: Generates and sends the daily plan.
- `Run_Evening_Reminder.bat`: Sends the evening status reminder.
