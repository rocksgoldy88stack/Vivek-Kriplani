import { site } from "@/content/site";

export default function FeaturedOn() {
  const { title, logos } = site.featuredOn;
  // Duplicate the list so the marquee can loop seamlessly.
  const loop = [...logos, ...logos];

  return (
    <section className="border-y border-border bg-surface py-10">
      <div className="mx-auto max-w-6xl px-5 sm:px-8">
        <p className="mb-6 text-center text-xs font-semibold uppercase tracking-[0.22em] text-muted">
          {title}
        </p>
        <div className="marquee-mask overflow-hidden">
          <div className="flex w-max animate-marquee items-center gap-12 sm:gap-16">
            {loop.map((logo, i) => (
              <span
                key={`${logo}-${i}`}
                className="whitespace-nowrap font-serif text-xl font-semibold text-muted/70 transition-colors hover:text-foreground sm:text-2xl"
              >
                {logo}
              </span>
            ))}
          </div>
        </div>
      </div>
    </section>
  );
}
