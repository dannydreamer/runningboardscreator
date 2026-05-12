# הוראות פריסה — Workshop Attendance Tool

## מה תצטרכי לפני שמתחילים
1. חשבון GitHub חינמי — https://github.com
2. חשבון Vercel חינמי — https://vercel.com
3. API Key של Anthropic — https://console.anthropic.com

---

## שלב 1: העלאה ל-GitHub

1. היכנסי ל-GitHub → לחצי על **"New repository"**
2. תני שם: `workshop-attendance`
3. השאירי אותו **Public** (או Private — שניהם עובדים)
4. לחצי **"Create repository"**
5. העלי את כל הקבצים מתיקיית הפרויקט (גרירה לדפדפן, או השתמשי ב-GitHub Desktop)

---

## שלב 2: פריסה ב-Vercel

1. היכנסי ל-https://vercel.com → **"Add New Project"**
2. חברי את חשבון GitHub שלך (בפעם הראשונה)
3. בחרי את הריפו `workshop-attendance` → לחצי **"Import"**
4. **חשוב:** לפני לחיצה על Deploy, פתחי **"Environment Variables"**
   - Name: `ANTHROPIC_API_KEY`
   - Value: המפתח שלך (מתחיל ב-`sk-ant-...`)
   - לחצי **"Add"**
5. לחצי **"Deploy"**
6. תוך כ-2 דקות תקבלי קישור כמו: `workshop-attendance.vercel.app`

---

## שלב 3: שיתוף

שלחי את הקישור לכל מי שצריך להשתמש בכלי. אין צורך בהתקנה, אין צורך ב-API key מצד המשתמשים.

---

## עדכונים בעתיד

אם תרצי לשנות משהו, ערכי את הקבצים ב-GitHub — Vercel תפרוס מחדש אוטומטית תוך דקה.

---

## עלויות

- GitHub: חינם
- Vercel: חינם (עד 100GB bandwidth בחודש — מספיק לשימוש סביר)
- Anthropic API: כ-0.003$ לכל קובץ תמונה שמעבדת. לשימוש רגיל — כמה דולרים בחודש בסך הכל.
