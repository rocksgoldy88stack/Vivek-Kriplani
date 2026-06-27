import Reveal from "./ui/Reveal";
import Icon from "./ui/Icon";
import { ArrowUpRight } from "lucide-react";
import { site } from "@/content/site";

export default function Connect() {
  const { title, lead, stats, socials } = site.connect;

  return (
    <section className="border-y border-border bg-surface">
      <div className="mx-auto max-w-6xl px-5 py-20 sm:px-8 sm:py-28">
        <div className="grid grid-cols-1 gap-12 lg:grid-cols-2 lg:gap-16">
          {/* Left: heading + stats */}
          <Reveal>
            <p className="mb-3 text-sm font-semibold uppercase tracking-[0.18em] text-accent">
              Connect
            </p>
            <h2 className="font-serif text-3xl font-semibold leading-tight tracking-tight sm:text-4xl">
              {title}
            </h2>
            <p className="mt-4 max-w-md text-lg leading-relaxed text-muted">
              {lead}
            </p>

            <dl className="mt-10 grid grid-cols-3 gap-4">
              {stats.map((s) => (
                <div key={s.label}>
                  <dt className="font-serif text-3xl font-bold text-accent sm:text-4xl">
                    {s.value}
                  </dt>
                  <dd className="mt-1 text-sm text-muted">{s.label}</dd>
                </div>
              ))}
            </dl>
          </Reveal>

          {/* Right: social links */}
          <div className="flex flex-col gap-3">
            {socials.map((s, i) => (
              <Reveal key={s.label} delay={i * 0.06}>
                <a
                  href={s.href}
                  target="_blank"
                  rel="noopener noreferrer"
                  className="group flex items-center justify-between rounded-xl border border-border bg-background px-5 py-4 transition-colors hover:border-accent"
                >
                  <span className="flex items-center gap-4">
                    <span className="grid h-11 w-11 place-items-center rounded-lg bg-accent-soft text-accent">
                      <Icon name={s.icon} className="h-5 w-5" />
                    </span>
                    <span>
                      <span className="block font-semibold">{s.label}</span>
                      {s.handle && (
                        <span className="block text-sm text-muted">
                          {s.handle}
                        </span>
                      )}
                    </span>
                  </span>
                  <ArrowUpRight className="h-5 w-5 text-muted transition-all group-hover:translate-x-0.5 group-hover:-translate-y-0.5 group-hover:text-accent" />
                </a>
              </Reveal>
            ))}
          </div>
        </div>
      </div>
    </section>
  );
}
