/**
 * ============================================================================
 *  SITE CONTENT — EDIT THIS FILE TO PERSONALIZE YOUR WEBSITE
 * ============================================================================
 *  Everything that appears on the site is driven from this single file.
 *  Replace the placeholder values below with your own details.
 *  No need to touch the React components unless you want to change layout.
 * ============================================================================
 */

export type NavLink = { label: string; href: string };

export type Role = {
  title: string;
  description: string;
  icon: "rocket" | "pen" | "book" | "trending" | "mic" | "code";
};

export type SocialLink = {
  label: string;
  href: string;
  handle?: string;
  icon:
    | "instagram"
    | "youtube"
    | "linkedin"
    | "twitter"
    | "github"
    | "globe"
    | "mail";
};

export type Venture = {
  name: string;
  tagline: string;
  description: string;
  href?: string;
  status?: string; // e.g. "Building", "Live", "Acquired"
};

export type Work = {
  title: string;
  category: string; // e.g. "Book", "Product", "Talk"
  description: string;
  href?: string;
  badge?: string; // e.g. "#1 Bestseller"
};

export type Post = {
  title: string;
  date: string; // human readable, e.g. "June 27, 2026"
  excerpt: string;
  href?: string;
};

export type Stat = { value: string; label: string };

export const site = {
  /* ---------------------------------------------------------------- META */
  meta: {
    name: "Vivek Vinod Kriplani",
    role: "Chartered Accountant · Risk & Tech Strategist",
    // Used for the browser tab + SEO
    title: "Vivek Vinod Kriplani — Chartered Accountant | Risk & Tech Strategist",
    description:
      "Vivek Vinod Kriplani is a Chartered Accountant specializing in corporate finance, internal audit and risk advisory — exploring how AI and Blockchain transform automated controls, audit trails and risk management.",
    // Absolute URL of the deployed site (used for SEO / Open Graph)
    url: "https://vivekkriplani.com",
    // Short word/initials used in the navbar logo
    logo: "VK",
  },

  /* ---------------------------------------------------------------- NAV */
  nav: [
    { label: "About", href: "#about" },
    { label: "Services", href: "#ventures" },
    { label: "Focus", href: "#work" },
    { label: "Insights", href: "#writing" },
    { label: "Contact", href: "#contact" },
  ] as NavLink[],

  /* ---------------------------------------------------------------- HERO */
  hero: {
    // The big philosophy/tagline statement
    headline: "Chartered Accountant by profession. Tech & risk strategist by passion.",
    subline:
      "I work at the intersection of corporate finance, internal audit and emerging technology — helping organizations stay compliant, resilient and ready for what's next.",
    // Short intro paragraph below the headline
    intro:
      "I'm Vivek — a CA who pairs financial rigour with a builder's curiosity for AI and Blockchain, turning complex compliance into a competitive advantage.",
    primaryCta: { label: "Work with me", href: "#contact" },
    secondaryCta: { label: "Read my insights", href: "#writing" },
    // Optional headshot in /public (e.g. "/me.jpg"). Leave "" to show initials.
    photo: "",
  },

  /* --------------------------------------------------------- FEATURED ON */
  featuredOn: {
    title: "Areas of specialization",
    // Just names render as clean text logos; swap for image logos later.
    logos: [
      "Internal Audit",
      "Risk Advisory",
      "Corporate Finance",
      "Ind AS / IFRS",
      "SOX Controls",
      "AI & Automation",
      "Blockchain",
      "GST & Compliance",
    ],
  },

  /* ----------------------------------------------------------- WHAT I DO */
  about: {
    title: "What I do",
    lead:
      "Chartered Accountant (CA) with a sharp focus on corporate finance, internal audit, and risk advisory. I specialize in navigating complex compliance landscapes while actively exploring how cutting-edge technologies like AI and Blockchain can revolutionize automated controls, audit trails, and risk management. I thrive at the intersection of numbers and innovation, helping organizations future-proof their operations against emerging digital risks.",
    roles: [
      {
        title: "Corporate Finance",
        icon: "trending",
        description:
          "Financial strategy, planning and analysis that turns numbers into clear, confident business decisions.",
      },
      {
        title: "Internal Audit",
        icon: "book",
        description:
          "Risk-based audits, process reviews and controls testing that strengthen governance and assurance.",
      },
      {
        title: "Risk Advisory",
        icon: "pen",
        description:
          "Identifying, assessing and mitigating enterprise risk before it becomes a problem — across finance, operations and compliance.",
      },
      {
        title: "Tech & Innovation",
        icon: "code",
        description:
          "Applying AI and Blockchain to automate controls, build tamper-proof audit trails and modernize risk management.",
      },
    ] as Role[],
  },

  /* ------------------------------------------------------------- CONNECT */
  connect: {
    title: "Let's connect",
    lead:
      "Open to advisory engagements, collaborations and conversations on where finance meets technology. Reach out — I read everything.",
    stats: [
      { value: "CA", label: "Chartered Accountant" },
      { value: "Finance", label: "Audit & Risk" },
      { value: "AI + Chain", label: "Tech-forward" },
    ] as Stat[],
    socials: [
      { label: "LinkedIn", href: "https://linkedin.com", handle: "Add your profile URL", icon: "linkedin" },
      { label: "X / Twitter", href: "https://x.com", handle: "Add your handle", icon: "twitter" },
      { label: "Email", href: "mailto:vivekkriplani88@gmail.com", handle: "vivekkriplani88@gmail.com", icon: "mail" },
    ] as SocialLink[],
  },

  /* ------------------------------------------------------------ VENTURES */
  ventures: {
    title: "Services",
    lead: "How I help organizations stay compliant, resilient and ahead of risk.",
    items: [
      {
        name: "Corporate Finance",
        tagline: "Strategy that scales",
        description:
          "Financial planning, analysis and structuring — from budgeting and forecasting to capital decisions and transaction support.",
      },
      {
        name: "Internal Audit",
        tagline: "Assurance you can trust",
        description:
          "Independent, risk-based internal audits and controls reviews that surface gaps and strengthen your governance framework.",
      },
      {
        name: "Risk Advisory",
        tagline: "See risk before it sees you",
        description:
          "Enterprise risk assessments, compliance frameworks and regulatory readiness tailored to your industry and scale.",
      },
    ] as Venture[],
  },

  /* ---------------------------------------------------------------- WORK */
  work: {
    title: "Where finance meets technology",
    lead: "The ideas and initiatives I'm most excited about right now.",
    items: [
      {
        title: "AI-Powered Audit Automation",
        category: "Focus Area",
        badge: "AI",
        description:
          "Using machine learning to automate audit testing, detect anomalies and move from sample-based checks to full-population assurance.",
      },
      {
        title: "Blockchain Audit Trails",
        category: "Focus Area",
        badge: "Blockchain",
        description:
          "Designing immutable, real-time verifiable audit trails that make tampering evident and reconciliation effortless.",
      },
      {
        title: "Automated Internal Controls",
        category: "Focus Area",
        description:
          "Continuous controls monitoring that replaces periodic checks with always-on assurance and early-warning alerts.",
      },
      {
        title: "Digital Risk Management",
        category: "Focus Area",
        description:
          "Future-proofing operations against emerging digital, cyber and regulatory risks with a technology-first risk posture.",
      },
    ] as Work[],
  },

  /* ------------------------------------------------------------- WRITING */
  writing: {
    title: "Insights",
    lead: "Notes on finance, audit, risk and the technology reshaping all three.",
    ctaLabel: "Read all articles",
    ctaHref: "#",
    posts: [
      {
        title: "Can AI replace the auditor?",
        date: "June 26, 2026",
        excerpt:
          "Short answer: no — but the auditor who uses AI will replace the one who doesn't. A practical look at where automation helps and where judgment still wins.",
        href: "#",
      },
      {
        title: "Blockchain and the future of the audit trail",
        date: "June 20, 2026",
        excerpt:
          "What changes when your ledger can't be quietly edited? A grounded take on immutable records, real-time assurance and what it means for trust.",
        href: "#",
      },
      {
        title: "Building internal controls that don't break",
        date: "June 14, 2026",
        excerpt:
          "Most controls fail not from bad design but from drift. Here's how continuous monitoring keeps them honest long after the audit is done.",
        href: "#",
      },
    ] as Post[],
  },

  /* ------------------------------------------------------------- CONTACT */
  contact: {
    title: "Let's work together",
    lead:
      "Whether it's an advisory engagement, a collaboration, or a conversation about finance and technology — I'd love to hear from you. Fill in the form below and I'll get back to you.",
    email: "vivekkriplani88@gmail.com",
    // The form posts here. Replace with your Formspree/Getform endpoint,
    // or leave "" to fall back to a mailto: link.
    formEndpoint: "",
  },

  /* -------------------------------------------------------------- FOOTER */
  footer: {
    note: "Where numbers meet innovation.",
  },
};

export type Site = typeof site;
