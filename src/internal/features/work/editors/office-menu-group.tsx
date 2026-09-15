import type { HTMLAttributes, ReactNode } from 'react';

/**
 * Labeled group for children of `role="menu"`.
 * `fieldset` is not a valid menu child; use `role="group"` instead.
 */
export function OfficeMenuGroup({
  ariaLabel,
  children,
  className,
  ...props
}: {
  ariaLabel: string;
  children: ReactNode;
  className?: string;
} & Omit<HTMLAttributes<HTMLDivElement>, 'role' | 'aria-label' | 'children'>) {
  return (
    // biome-ignore lint/a11y/useSemanticElements: role="menu" forbids fieldset; group scopes menuitem* sets.
    <div role="group" className={className} aria-label={ariaLabel} {...props}>
      {children}
    </div>
  );
}
