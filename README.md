# Office Tabs - Workspace Manager for Microsoft Office

מערכת כרטיסיות (Tabs) מתקדמת ומודרנית עבור Microsoft Word, Excel ו-PowerPoint, המיועדת לסביבת Windows 10/11 וכתובה ב-C# וב-.NET 8.

## תכונות עיקריות
- **זיהוי ועיגון אוטומטי**: מזהה אוטומטית פתיחה של מסמכי Word, Excel ו-PowerPoint, ומעביר אותם לממשק כרטיסיות מרכזי בשימוש ב-Windows API.
- **Fluent UI**: עיצוב נקי בסגנון Fluent Design עם תמיכה מובנית בערכות נושא (מצב כהה / בהיר) [1].
- **DPI-Aware**: התאמה מושלמת למסכים ברזולוציות גבוהות ומערכות מרובות מסכים.
- **הפעלה ברקע (System Tray)**: ריצה שקטה ברקע עם תפריט ניהול מהיר ב-System Tray של Windows.
- **ללא תלות ב-COM Add-ins**: אינו מאט את עליית היישומים של Office ופועל ללא צורך בהרשאות מנהל מערכת (Non-Admin Installer).

## דרישות מערכת
- Windows 10 / Windows 11
- Microsoft 365, Office 2021 או Office 2024
- .NET 8.0 Runtime

## ארכיטקטורה וטכנולוגיה
הפרויקט אינו משתמש ב-Add-ins מסורתיים של Office, מאחר ואלו נוטים לקרוס, להיות מנוטרלים על ידי מדיניות אבטחה, ולהשפיע לרעה על זמני הטעינה. במקום זאת, מומש מנגנון המבוסס על **WinEventHooks** גלובליים המזהה יצירת חלונות חדשים במערכת השייכים ל-Classes הידועים של Office:
1. `OpusApp` - Microsoft Word
2. `XLMAIN` - Microsoft Excel
3. `PPTFrameClass` - Microsoft PowerPoint

עם זיהוי חלון כזה, המערכת משתמשת בפעולת `SetParent` של מערכת ההפעלה כדי להטמיע את היישום בתוך טופס הניהול (Tab Container) שלה [1], מסירה את גבולות החלון המקוריים על ידי מניפולציה של ה-Window Style ומציגה אותו ככרטיסייה ייעודית במסך הראשי.

## כיצד לבנות ולהריץ את הפרויקט ידנית

1. ודא שברשותך .NET 8 SDK מותקן במחשב.
2. שכפל את ה-Repository:
   ```bash
   git clone https://github.com/your-username/OfficeTabs.git
   cd OfficeTabs
