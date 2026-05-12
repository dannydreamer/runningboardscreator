export const metadata = { title: 'נוכחות סדנאות', description: 'כלי לניהול נוכחות בסדנאות' };

export default function RootLayout({ children }) {
  return (
    <html lang="he" dir="rtl">
      <body style={{ margin: 0 }}>{children}</body>
    </html>
  );
}
