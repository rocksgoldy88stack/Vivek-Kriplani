"use client";

import { useState, type FormEvent } from "react";
import { Mail, Send, CheckCircle2 } from "lucide-react";
import Reveal from "./ui/Reveal";
import { site } from "@/content/site";

type Status = "idle" | "submitting" | "success" | "error";

export default function Contact() {
  const { title, lead, email, formEndpoint } = site.contact;
  const [status, setStatus] = useState<Status>("idle");

  async function handleSubmit(e: FormEvent<HTMLFormElement>) {
    e.preventDefault();
    const form = e.currentTarget;
    const data = new FormData(form);

    // No backend configured: fall back to opening the user's mail client.
    if (!formEndpoint) {
      const name = String(data.get("name") || "");
      const from = String(data.get("email") || "");
      const message = String(data.get("message") || "");
      const subject = encodeURIComponent(`Inquiry from ${name}`);
      const body = encodeURIComponent(`${message}\n\n— ${name} (${from})`);
      window.location.href = `mailto:${email}?subject=${subject}&body=${body}`;
      return;
    }

    try {
      setStatus("submitting");
      const res = await fetch(formEndpoint, {
        method: "POST",
        headers: { Accept: "application/json" },
        body: data,
      });
      if (!res.ok) throw new Error("Request failed");
      setStatus("success");
      form.reset();
    } catch {
      setStatus("error");
    }
  }

  const fieldClass =
    "w-full rounded-xl border border-border bg-background px-4 py-3 outline-none transition-colors placeholder:text-muted/70 focus:border-accent focus:ring-2 focus:ring-accent/20";

  return (
    <section id="contact" className="border-t border-border bg-surface">
      <div className="mx-auto grid max-w-6xl grid-cols-1 gap-12 px-5 py-20 sm:px-8 sm:py-28 lg:grid-cols-2 lg:gap-16">
        <Reveal>
          <p className="mb-3 text-sm font-semibold uppercase tracking-[0.18em] text-accent">
            Contact
          </p>
          <h2 className="font-serif text-3xl font-semibold leading-tight tracking-tight sm:text-4xl">
            {title}
          </h2>
          <p className="mt-4 max-w-md text-lg leading-relaxed text-muted">
            {lead}
          </p>
          <a
            href={`mailto:${email}`}
            className="mt-8 inline-flex items-center gap-3 rounded-xl border border-border bg-background px-5 py-4 transition-colors hover:border-accent"
          >
            <span className="grid h-11 w-11 place-items-center rounded-lg bg-accent-soft text-accent">
              <Mail className="h-5 w-5" />
            </span>
            <span>
              <span className="block text-sm text-muted">Email me directly</span>
              <span className="block font-semibold">{email}</span>
            </span>
          </a>
        </Reveal>

        <Reveal delay={0.1}>
          {status === "success" ? (
            <div className="flex h-full flex-col items-center justify-center rounded-2xl border border-border bg-background p-10 text-center">
              <CheckCircle2 className="h-12 w-12 text-accent" />
              <h3 className="mt-4 font-serif text-2xl font-semibold">
                Thank you!
              </h3>
              <p className="mt-2 text-muted">
                Your message is on its way. I&apos;ll get back to you soon.
              </p>
            </div>
          ) : (
            <form
              onSubmit={handleSubmit}
              className="rounded-2xl border border-border bg-background p-6 sm:p-8"
            >
              <div className="grid grid-cols-1 gap-4 sm:grid-cols-2">
                <div className="sm:col-span-1">
                  <label htmlFor="name" className="mb-1.5 block text-sm font-medium">
                    Name
                  </label>
                  <input id="name" name="name" required className={fieldClass} placeholder="Jane Doe" />
                </div>
                <div className="sm:col-span-1">
                  <label htmlFor="email" className="mb-1.5 block text-sm font-medium">
                    Email
                  </label>
                  <input id="email" name="email" type="email" required className={fieldClass} placeholder="jane@company.com" />
                </div>
                <div className="sm:col-span-2">
                  <label htmlFor="subject" className="mb-1.5 block text-sm font-medium">
                    Subject
                  </label>
                  <input id="subject" name="subject" className={fieldClass} placeholder="Speaking engagement / collaboration" />
                </div>
                <div className="sm:col-span-2">
                  <label htmlFor="message" className="mb-1.5 block text-sm font-medium">
                    Message
                  </label>
                  <textarea id="message" name="message" required rows={5} className={`${fieldClass} resize-none`} placeholder="Tell me a little about what you have in mind..." />
                </div>
              </div>

              {status === "error" && (
                <p className="mt-4 text-sm text-red-500">
                  Something went wrong. Please try again or email me directly.
                </p>
              )}

              <button
                type="submit"
                disabled={status === "submitting"}
                className="group mt-6 inline-flex w-full items-center justify-center gap-2 rounded-xl bg-accent px-6 py-3.5 font-semibold text-white transition-transform hover:scale-[1.01] disabled:opacity-60 sm:w-auto"
              >
                {status === "submitting" ? "Sending..." : "Send message"}
                <Send className="h-4 w-4 transition-transform group-hover:translate-x-0.5" />
              </button>
            </form>
          )}
        </Reveal>
      </div>
    </section>
  );
}
