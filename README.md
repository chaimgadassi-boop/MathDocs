# MathDocs Local · מערכת מתמטיקה מקומית ל־Windows

**גרסת תשתית 0.1 – אבטיפוס פעיל, לא מוצר סופי.**

מטרה: לייצר מסמכי Word בעברית עם נוסחאות, גבולות ומטריצות מקוריים של Word (OMML).
הם אינם תמונות ולכן ניתנים לעריכה והעתקה בין מסמכי Word.
גרפים וחישובים אינטראקטיביים נוצרים באמצעות Excel המותקן ב־Windows.

## התשתית כבר כוללת
- מחולל DOCX בשפה העברית עם כיוון ימין־לשמאל.
- נוסחאות Word מקוריות (שברים, חזקות, גבולות ומטריצות).
- חוזה נתונים JSON, כך שאפשר להכין תוכן כאן בצ'אט ולייבא אותו לתוכנה.
- דוגמת Excel עם מטריצה, דטרמיננטה וגרף מבוסס תאים (דורש Excel ב־Windows).
- בדיקות אוטומטיות, קובץ הנחיות Codex ותכנון ארכיטקטורה.
- OpenAI API *אופציונלי בלבד*, כבוי כברירת מחדל. אין שימוש ב־Claude.

## הפעלה ב־Windows
1. התקן Python 3.10+ ואת Microsoft Word ו־Excel.
2. חלץ את הפרויקט לתיקייה.
3. פתח PowerShell בתוך התיקייה.
4. הרץ:

```powershell
py -m venv .venv
.\\.venv\\Scripts\\Activate.ps1
python -m pip install -e ".[dev,windows]"
python -m pytest -q
python -m mathdocs demo --output output\\example.docx
```

**פתח ב־Word את `output/example.docx`** ונסה לערוך את הנוסחה ואת המטריצה.
קובץ חדש מסיכום אחר:

```powershell
python -m mathdocs build examples\\lesson_limits.json --output output\\limits.docx
```

Excel (Windows עם Excel מותקן):

```powershell
python -m mathdocs excel-demo --output output\\matrix_chart.xlsx
```

## עבודה עם Codex
התקן Codex CLI לפי האתר הרשמי של OpenAI ופתח אותו בשורש הפרויקט:

```powershell
codex
```

Codex יקרא את `AGENTS.md` ואת `docs/ARCHITECTURE.md`. אין צורך ב־Claude.

## חשוב לדעת
- **זה אבטיפוס**: הנוסחאות נבנות ב־OMML, אבל בדיקת Word אמיתי (פתיחה, עריכה, העתקה) צריכה להתבצע על מחשב Windows שלך.
- Microsoft Word עשוי לשנות גופנים ועיצוב בהעתקה למסמך עם סגנונות שונים.
- Excel COM לא פועל ב־Linux או בסביבת ענן ללא Office.
- אין סיבה לרכוש מפתח API כדי להריץ את הדגמה.
- להתקנת Codex אפשר להתחבר עם ChatGPT, בהתאם לתוכנית ולמגבלות השימוש.
- טקסט מחומר לימוד לא עובר לענן בשלב יצירת Word.

## מבנה תיקיות
```
src/mathdocs/      מנועי קבצים, OMML, API אופציונלי ו־CLI
examples/          שיעור לדוגמה ב־JSON
tests/             בדיקות מבנה וכשלים
docs/              ארכיטקטורה ומשימות להמשך
AGENTS.md          כללי עבודה ל־Codex
```
