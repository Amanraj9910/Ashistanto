import { cn } from '@/lib/utils';

type BrandLogoProps = {
  className?: string;
  imageClassName?: string;
  showText?: boolean;
};

export function BrandLogo({ className, imageClassName, showText = true }: BrandLogoProps) {
  return (
    <div className={cn('flex items-center gap-3', className)}>
      <img
        src="/img/Hosho DIgital-Logo.jpg"
        alt="Ashistanto"
        className={cn('h-10 w-auto object-contain', imageClassName)}
      />
      {showText && <span className="sr-only">ASHISTANTO</span>}
    </div>
  );
}
