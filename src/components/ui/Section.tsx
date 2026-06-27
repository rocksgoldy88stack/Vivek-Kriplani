import type { ReactNode } from "react";
import Reveal from "./Reveal";

type SectionProps = {
  id?: string;
  eyebrow?: string;
  title?: string;
  lead?: string;
  children: ReactNode;
  className?: string;
};

/** Standard section shell with an optional eyebrow / title / lead header. */
export default function Section({
  id,
  eyebrow,
  title,
  lead,
  children,
  className = "",
}: SectionProps) {
  return (
    <section
      id={id}
      className={`mx-auto w-full max-w-6xl px-5 py-20 sm:px-8 sm:py-28 ${className}`}
    >
      {(eyebrow || title || lead) && (
        <Reveal className="mb-12 max-w-2xl">
          {eyebrow && (
            <p className="mb-3 text-sm font-semibold uppercase tracking-[0.18em] text-accent">
              {eyebrow}
            </p>
          )}
          {title && (
            <h2 className="font-serif text-3xl font-semibold leading-tight tracking-tight sm:text-4xl">
              {title}
            </h2>
          )}
          {lead && (
            <p className="mt-4 text-lg leading-relaxed text-muted">{lead}</p>
          )}
        </Reveal>
      )}
      {children}
    </section>
  );
}
