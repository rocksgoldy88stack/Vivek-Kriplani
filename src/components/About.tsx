import Section from "./ui/Section";
import Reveal from "./ui/Reveal";
import Icon from "./ui/Icon";
import { site } from "@/content/site";

export default function About() {
  const { title, lead, roles } = site.about;

  return (
    <Section id="about" eyebrow="About" title={title} lead={lead}>
      <div className="grid grid-cols-1 gap-5 sm:grid-cols-2">
        {roles.map((role, i) => (
          <Reveal key={role.title} delay={i * 0.08}>
            <article className="group h-full rounded-2xl border border-border bg-surface p-7 transition-all hover:-translate-y-1 hover:border-accent/40 hover:shadow-lg hover:shadow-accent/5">
              <div className="mb-5 grid h-12 w-12 place-items-center rounded-xl bg-accent-soft text-accent transition-colors group-hover:bg-accent group-hover:text-white">
                <Icon name={role.icon} className="h-6 w-6" />
              </div>
              <h3 className="font-serif text-xl font-semibold">{role.title}</h3>
              <p className="mt-3 leading-relaxed text-muted">
                {role.description}
              </p>
            </article>
          </Reveal>
        ))}
      </div>
    </Section>
  );
}
