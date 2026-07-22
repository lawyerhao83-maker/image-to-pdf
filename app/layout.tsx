import type { Metadata } from "next";
import { Geist, Geist_Mono } from "next/font/google";
import { headers } from "next/headers";
import "./globals.css";

const geistSans = Geist({
  variable: "--font-geist-sans",
  subsets: ["latin"],
});

const geistMono = Geist_Mono({
  variable: "--font-geist-mono",
  subsets: ["latin"],
});

export async function generateMetadata(): Promise<Metadata> {
  const requestHeaders = await headers();
  const host = requestHeaders.get("x-forwarded-host") ?? requestHeaders.get("host") ?? "localhost:3000";
  const protocol = requestHeaders.get("x-forwarded-proto") ?? (host.startsWith("localhost") ? "http" : "https");
  const ogImage = `${protocol}://${host}/og.png`;

  return {
    title: "PictoPDF｜JPG 圖片轉 PDF",
    description: "拖放 JPG、自由排序並在瀏覽器中安全轉換成 PDF。",
    icons: { icon: "/favicon.svg", shortcut: "/favicon.svg" },
    openGraph: {
      title: "PictoPDF｜圖片排好，一鍵成冊。",
      description: "拖放、排序、旋轉 JPG，在瀏覽器中安全轉成 PDF。",
      type: "website",
      images: [{ url: ogImage, width: 1536, height: 1024, alt: "PictoPDF JPG 轉 PDF 工具" }],
    },
    twitter: {
      card: "summary_large_image",
      title: "PictoPDF｜圖片排好，一鍵成冊。",
      description: "拖放、排序、旋轉 JPG，在瀏覽器中安全轉成 PDF。",
      images: [ogImage],
    },
  };
}

export default function RootLayout({
  children,
}: Readonly<{
  children: React.ReactNode;
}>) {
  return (
    <html lang="zh-Hant">
      <body
        className={`${geistSans.variable} ${geistMono.variable} antialiased`}
      >
        {children}
      </body>
    </html>
  );
}
