import type { Metadata } from "next";

export const metadata: Metadata = {
  title: "Generate Excel",
  description: "Generated DArt",
};

export default function RootLayout({
  children,
}: Readonly<{
  children: React.ReactNode;
}>) {
  return (
    <html lang="en">
      <body>
        {children}
      </body>
    </html>
  );
}
