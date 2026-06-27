"use client";

import { motion } from "framer-motion";
import { ArrowRight } from "lucide-react";
import { site } from "@/content/site";

const container = {
  hidden: {},
  show: { transition: { staggerChildren: 0.12, delayChildren: 0.1 } },
};
const item = {
  hidden: { opacity: 0, y: 22 },
  show: { opacity: 1, y: 0, transition: { duration: 0.7, ease: [0.22, 1, 0.36, 1] as const } },
};

export default function Hero() {
  const { hero, meta } = site;

  return (
    <section
      id="top"
      className="relative overflow-hidden bg-dotted pt-32 pb-20 sm:pt-40 sm:pb-28"
    >
      {/* soft accent glow */}
      <div
        aria-hidden
        className="pointer-events-none absolute -top-24 left-1/2 h-[420px] w-[820px] -translate-x-1/2 rounded-full opacity-60 blur-3xl"
        style={{
          background:
            "radial-gradient(closest-side, var(--accent-soft), transparent)",
        }}
      />

      <motion.div
        variants={container}
        initial="hidden"
        animate="show"
        className="relative mx-auto grid max-w-6xl grid-cols-1 items-center gap-12 px-5 sm:px-8 lg:grid-cols-[1.4fr_1fr]"
      >
        <div>
          <motion.span
            variants={item}
            className="inline-flex items-center gap-2 rounded-full border border-border bg-surface px-4 py-1.5 text-sm font-medium text-muted"
          >
            <span className="h-2 w-2 rounded-full bg-accent" />
            {meta.role}
          </motion.span>

          <motion.h1
            variants={item}
            className="mt-6 font-serif text-5xl font-semibold leading-[1.05] tracking-tight text-balance sm:text-6xl lg:text-7xl"
          >
            {hero.headline}
          </motion.h1>

          <motion.p
            variants={item}
            className="mt-6 max-w-xl text-xl leading-relaxed text-muted text-balance"
          >
            {hero.subline}
          </motion.p>

          <motion.p
            variants={item}
            className="mt-5 max-w-xl leading-relaxed text-muted"
          >
            {hero.intro}
          </motion.p>

          <motion.div variants={item} className="mt-9 flex flex-wrap gap-3">
            <a
              href={hero.primaryCta.href}
              className="group inline-flex items-center gap-2 rounded-full bg-accent px-6 py-3 font-semibold text-white transition-transform hover:scale-[1.02]"
            >
              {hero.primaryCta.label}
              <ArrowRight className="h-4 w-4 transition-transform group-hover:translate-x-1" />
            </a>
            <a
              href={hero.secondaryCta.href}
              className="inline-flex items-center gap-2 rounded-full border border-border bg-surface px-6 py-3 font-semibold transition-colors hover:border-accent hover:text-accent"
            >
              {hero.secondaryCta.label}
            </a>
          </motion.div>
        </div>

        {/* Portrait / initials card */}
        <motion.div
          variants={item}
          className="relative mx-auto w-full max-w-sm lg:mx-0"
        >
          <div className="absolute -inset-3 -rotate-3 rounded-[2rem] bg-accent-soft" />
          <div className="relative aspect-[4/5] overflow-hidden rounded-[1.75rem] border border-border bg-surface shadow-sm">
            {hero.photo ? (
              // eslint-disable-next-line @next/next/no-img-element
              <img
                src={hero.photo}
                alt={meta.name}
                className="h-full w-full object-cover"
              />
            ) : (
              <div className="flex h-full w-full flex-col items-center justify-center gap-4">
                <span className="grid h-28 w-28 place-items-center rounded-full bg-foreground font-serif text-4xl font-bold text-background">
                  {meta.logo}
                </span>
                <p className="px-6 text-center text-sm text-muted">
                  Add a photo by setting{" "}
                  <code className="rounded bg-accent-soft px-1.5 py-0.5 text-accent">
                    hero.photo
                  </code>{" "}
                  in <code className="text-accent">site.ts</code>
                </p>
              </div>
            )}
          </div>
        </motion.div>
      </motion.div>
    </section>
  );
}
