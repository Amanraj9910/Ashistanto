import type { Metadata } from 'next';
import { Manrope, Poppins } from 'next/font/google';
import './globals.css';
import { Providers } from './providers';

const poppins = Poppins({
  subsets: ['latin'],
  weight: ['400', '500', '600', '700', '800'],
  display: 'swap',
  variable: '--font-poppins'
});

const manrope = Manrope({
  subsets: ['latin'],
  weight: ['400', '500', '600', '700', '800'],
  display: 'swap',
  variable: '--font-manrope'
});

export const metadata: Metadata = {
  metadataBase: new URL('https://ashistanto.com'),
  title: {
    default: 'Ashistanto - AI Assistant Workspace',
    template: '%s | Ashistanto'
  },
  description: 'Ashistanto is an AI assistant workspace for Microsoft 365 workflows, email, meetings, files, and enterprise actions.',
  applicationName: 'Ashistanto',
  authors: [{ name: 'Hosho Digital' }],
  keywords: ['Ashistanto', 'AI assistant', 'Microsoft 365', 'email assistant', 'enterprise AI', 'workflow automation'],
  alternates: {
    canonical: '/'
  },
  openGraph: {
    title: 'Ashistanto - AI Assistant Workspace',
    description: 'Manage emails, meetings, files, and enterprise actions through an AI assistant workspace.',
    url: '/',
    siteName: 'Ashistanto',
    images: [
      {
        url: '/img/Hosho DIgital-Logo.jpg',
        width: 1200,
        height: 630,
        alt: 'Ashistanto by Hosho Digital'
      }
    ],
    locale: 'en_US',
    type: 'website'
  },
  twitter: {
    card: 'summary_large_image',
    title: 'Ashistanto - AI Assistant Workspace',
    description: 'An AI assistant workspace for Microsoft 365 workflows.',
    images: ['/img/Hosho DIgital-Logo.jpg']
  },
  icons: {
    icon: '/img/favicon.ico.png',
    shortcut: '/img/favicon.ico.png',
    apple: '/img/favicon.ico.png'
  }
};

export default function RootLayout({ children }: Readonly<{ children: React.ReactNode }>) {
  return (
    <html lang="en" suppressHydrationWarning>
      <body className={`${poppins.variable} ${manrope.variable}`}>
        <Providers>{children}</Providers>
      </body>
    </html>
  );
}
