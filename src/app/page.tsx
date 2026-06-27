import Navbar from "@/components/Navbar";
import Hero from "@/components/Hero";
import FeaturedOn from "@/components/FeaturedOn";
import About from "@/components/About";
import Connect from "@/components/Connect";
import Ventures from "@/components/Ventures";
import Work from "@/components/Work";
import Writing from "@/components/Writing";
import Contact from "@/components/Contact";
import Footer from "@/components/Footer";

export default function Home() {
  return (
    <>
      <Navbar />
      <main className="flex-1">
        <Hero />
        <FeaturedOn />
        <About />
        <Ventures />
        <Work />
        <Connect />
        <Writing />
        <Contact />
      </main>
      <Footer />
    </>
  );
}
