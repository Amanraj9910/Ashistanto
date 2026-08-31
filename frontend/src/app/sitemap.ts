import type { MetadataRoute } from 'next';

// Required by `output: 'export'` — metadata routes must opt in to being prerendered.
export const dynamic = 'force-static';

// No `lastModified` here: under a static export it would freeze to build time and then
// misreport every page as modified on each deploy.
export default function sitemap(): MetadataRoute.Sitemap {
  return [
    { url: 'https://ashistanto.com/', changeFrequency: 'weekly', priority: 1 },
    { url: 'https://ashistanto.com/login', changeFrequency: 'monthly', priority: 0.5 },
    { url: 'https://ashistanto.com/chat', changeFrequency: 'weekly', priority: 0.8 }
  ];
}
