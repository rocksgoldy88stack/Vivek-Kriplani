import Section from "./ui/Section";
import Reveal from "./ui/Reveal";
import { ArrowRight, ArrowUpRight } from "lucide-react";
import { site } from "@/content/site";

export default function Writing() {
  const { title, lead, posts, ctaLabel, ctaHref } = site.writing;

  return (
    <Section id="writing" eyebrow="Insights" title={title} lead={lead}>
      <div className="grid grid-cols-1 gap-px overflow-hidden rounded-2xl border border-border bg-border">
        {posts.map((p, i) => {
          const inner = (
            <article className="flex flex-col gap-3 bg-surface p-7 transition-colors hover:bg-accent-soft/40 sm:flex-row sm:items-center sm:justify-between sm:gap-8">
              <div className="max-w-2xl">
                <div className="flex items-center gap-3">
                  <time className="text-xs font-semibold uppercase tracking-[0.14em] text-muted">
                    {p.date}
                  </time>
                </div>
                <h3 className="mt-2 font-serif text-xl font-semibold transition-colors group-hover:text-accent">
                  {p.title}
                </h3>
                <p className="mt-2 leading-relaxed text-muted">{p.excerpt}</p>
              </div>
              <ArrowUpRight className="hidden h-6 w-6 shrink-0 text-muted transition-all group-hover:translate-x-0.5 group-hover:-translate-y-0.5 group-hover:text-accent sm:block" />
            </article>
          );

          return (
            <Reveal key={p.title} delay={i * 0.05} className="group">
              {p.href ? (
                <a href={p.href} target="_blank" rel="noopener noreferrer" className="block">
                  {inner}
                </a>
              ) : (
                inner
              )}
            </Reveal>
          );
        })}
      </div>

      {ctaHref && (
        <Reveal className="mt-8">
          <a
            href={ctaHref}
            className="group inline-flex items-center gap-2 font-semibold text-accent"
          >
            {ctaLabel}
            <ArrowRight className="h-4 w-4 transition-transform group-hover:translate-x-1" />
          </a>
        </Reveal>
      )}
    </Section>
  );
}
