import {
  Rocket,
  PenLine,
  BookOpen,
  TrendingUp,
  Mic,
  Code2,
  Globe,
  Mail,
  type LucideIcon,
} from "lucide-react";
import {
  FaInstagram,
  FaYoutube,
  FaLinkedinIn,
  FaXTwitter,
  FaGithub,
} from "react-icons/fa6";
import type { IconType } from "react-icons";

type AnyIcon = LucideIcon | IconType;

const map: Record<string, AnyIcon> = {
  // roles (lucide)
  rocket: Rocket,
  pen: PenLine,
  book: BookOpen,
  trending: TrendingUp,
  mic: Mic,
  code: Code2,
  globe: Globe,
  mail: Mail,
  // brand / social (react-icons)
  instagram: FaInstagram,
  youtube: FaYoutube,
  linkedin: FaLinkedinIn,
  twitter: FaXTwitter,
  github: FaGithub,
};

export default function Icon({
  name,
  className,
}: {
  name: string;
  className?: string;
}) {
  const Cmp = map[name] ?? Globe;
  return <Cmp className={className} aria-hidden />;
}
