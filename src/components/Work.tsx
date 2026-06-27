import Section from "./ui/Section";
import Reveal from "./ui/Reveal";
import { ArrowUpRight } from "lucide-react";
import { site } from "@/content/site";

export default function Work() {
  const { title, lead, items } = site.work;

  return (
    <Section id="work" eyebrow="Focus" title={title} lead={lead}>
      <div className="grid grid-cols-1 gap-5 sm:grid-cols-2">
        {items.map((w, i) => {
          const inner = (
            <article className="flex h-full gap-5 rounded-2xl border border-border bg-surface p-6 transition-all hover:-translate-y-1 hover:border-accent/40 hover:shadow-lg hover:shadow-accent/5">
              {/* spine/cover accent */}
              <div className="hidden w-1.5 shrink-0 rounded-full bg-gradient-to-b from-accent to-accent/30 sm:block" />
              <div className="flex flex-1 flex-col">
                <div className="flex items-center gap-2">
                  <span className="text-xs font-semibold uppercase tracking-[0.16em] text-muted">
                    {w.category}
                  </span>
                  {w.badge && (
                    <span className="rounded-full bg-accent-soft px-2.5 py-0.5 text-xs font-semibold text-accent">
                      {w.badge}
                    </span>
                  )}
                </div>
                <div className="mt-2 flex items-start justify-between gap-3">
                  <h3 className="font-serif text-xl font-semibold">{w.title}</h3>
                  {w.href && (
                    <ArrowUpRight className="h-5 w-5 shrink-0 text-muted transition-all group-hover:translate-x-0.5 group-hover:-translate-y-0.5 group-hover:text-accent" />
                  )}
                </div>
                <p className="mt-2 leading-relaxed text-muted">{w.description}</p>
              </div>
            </article>
          );

          return (
            <Reveal key={w.title} delay={i * 0.06} className="group h-full">
              {w.href ? (
                <a href={w.href} target="_blank" rel="noopener noreferrer" className="block h-full">
                  {inner}
                </a>
              ) : (
                inner
              )}
            </Reveal>
          );
        })}
      </div>
    </Section>
  );
}
