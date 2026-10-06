import * as React from 'react';

export interface IArrowIconProps {
  direction: 'left' | 'right';
}

/**
 * Pagination arrow drawn as an SVG line icon instead of a ←/→ text character:
 * a text glyph shrinks to a few unreadable pixels in a small circle (and sits
 * differently in every font), while this scales with its circle — 60% of the
 * button — and stays a crisp arrow at any size.
 */
export default function ArrowIcon(props: IArrowIconProps): React.ReactElement {
  const path = props.direction === 'left' ? 'M13.5 8H2.5M7 3.5L2.5 8L7 12.5' : 'M2.5 8H13.5M9 3.5L13.5 8L9 12.5';

  return (
    <svg
      viewBox="0 0 16 16"
      width="60%"
      height="60%"
      aria-hidden="true"
      focusable="false"
      style={{ display: 'block', flexShrink: 0 }}
    >
      <path d={path} fill="none" stroke="currentColor" strokeWidth="1.6" strokeLinecap="round" strokeLinejoin="round" />
    </svg>
  );
}
