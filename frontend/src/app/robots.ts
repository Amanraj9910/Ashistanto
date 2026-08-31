import type { MetadataRoute } from 'next';

// Required by `output: 'export'` — metadata routes must opt in to being prerendered.
export const dynamic = 'force-static';

export default function robots(): MetadataRoute.Robots {
  return {
    rules: {
      userAgent: '*',
      allow: '/'
    },
    sitemap: 'https://ashistanto.com/sitemap.xml'
  };
}
