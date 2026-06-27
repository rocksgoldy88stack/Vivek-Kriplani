import Section from "./ui/Section";
import Reveal from "./ui/Reveal";
import { ArrowUpRight } from "lucide-react";
import { site } from "@/content/site";

export default function Ventures() {
  const { title, lead, items } = site.ventures;

  return (
    <Section id="ventures" eyebrow="Services" title={title} lead={lead}>
      <div className="grid grid-cols-1 gap-5 md:grid-cols-3">
        {items.map((v, i) => {
          const inner = (
            <article className="flex h-full flex-col rounded-2xl border border-border bg-surface p-7 transition-all hover:-translate-y-1 hover:border-accent/40 hover:shadow-lg hover:shadow-accent/5">
              <div className="flex items-start justify-between gap-3">
                <h3 className="font-serif text-2xl font-semibold">{v.name}</h3>
                {v.href && (
                  <ArrowUpRight className="h-5 w-5 shrink-0 text-muted transition-all group-hover:translate-x-0.5 group-hover:-translate-y-0.5 group-hover:text-accent" />
                )}
              </div>
              {v.status && (
                <span className="mt-3 inline-flex w-fit items-center gap-1.5 rounded-full bg-accent-soft px-3 py-1 text-xs font-semibold text-accent">
                  <span className="h-1.5 w-1.5 rounded-full bg-accent" />
                  {v.status}
                </span>
              )}
              <p className="mt-4 font-medium">{v.tagline}</p>
              <p className="mt-2 leading-relaxed text-muted">{v.description}</p>
            </article>
          );

          return (
            <Reveal key={v.name} delay={i * 0.08} className="group h-full">
              {v.href ? (
                <a href={v.href} target="_blank" rel="noopener noreferrer" className="block h-full">
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
