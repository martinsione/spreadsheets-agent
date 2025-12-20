import { AnnouncementBadge } from "@/components/elements/announcement-badge";
import { ButtonLink, PlainButtonLink } from "@/components/elements/button";
import { Main } from "@/components/elements/main";
import { Screenshot } from "@/components/elements/screenshot";
import { ArrowNarrowRightIcon } from "@/components/icons/arrow-narrow-right-icon";
import { GitHubIcon } from "@/components/icons/social/github-icon";
import { XIcon } from "@/components/icons/social/x-icon";
import {
  FAQsTwoColumnAccordion,
  Faq,
} from "@/components/sections/faqs-two-column-accordion";
import {
  FooterCategory,
  FooterLink,
  FooterWithNewsletterFormCategoriesAndSocialIcons,
  SocialLink,
} from "@/components/sections/footer-with-newsletter-form-categories-and-social-icons";
import { HeroLeftAlignedWithDemo } from "@/components/sections/hero-left-aligned-with-demo";
import { NavbarLogo } from "@/components/sections/navbar-with-links-actions-and-centered-logo";

export default function Page() {
  return (
    <>
      {/* <NavbarWithLinksActionsAndCenteredLogo
        id="navbar"
        links={undefined}
        logo={
          <NavbarLogo href="#">
            <p className="font-serif text-2xl">OpenSheets</p>
          </NavbarLogo>
        }
        actions={undefined}
      /> */}
      <div className="pt-6 pl-6">
        <NavbarLogo href="#">
          <p className="font-serif text-2xl">OpenSheets</p>
        </NavbarLogo>
      </div>

      <Main>
        {/* Hero */}
        <HeroLeftAlignedWithDemo
          id="hero"
          eyebrow={
            <AnnouncementBadge
              href="https://github.com/martinsione/opensheets"
              text="OpenSheets is open source"
              cta="Star on GitHub"
            />
          }
          headline="The open source AI agent for spreadsheets."
          subheadline={
            <p>It lets you use AI to automate your spreadsheet tasks.</p>
          }
          cta={
            <div className="flex items-center gap-4">
              <ButtonLink href="#" size="lg">
                Install on Google Sheets
              </ButtonLink>

              <PlainButtonLink href="#" size="lg">
                See how it works <ArrowNarrowRightIcon />
              </PlainButtonLink>
            </div>
          }
          demo={
            <>
              <Screenshot
                className="rounded-md lg:hidden"
                wallpaper="green"
                placement="bottom-right"
              >
                {/** biome-ignore lint/performance/noImgElement: <> */}
                <img
                  src="/screenshot.png"
                  alt="OpenSheets for Google Sheets"
                  width={2000}
                  height={1408}
                  className="bg-white/75 max-md:hidden dark:hidden"
                />
              </Screenshot>
              <Screenshot
                className="rounded-lg max-lg:hidden"
                wallpaper="green"
                placement="bottom"
              >
                {/** biome-ignore lint/performance/noImgElement: <> */}
                <img
                  src="/screenshot.png"
                  alt="OpenSheets for Google Sheets"
                  className="bg-white/75"
                  width={3440}
                  height={1990}
                />
              </Screenshot>
            </>
          }
          footer={undefined}
        />

        {/* FAQs */}
        <FAQsTwoColumnAccordion id="faqs" headline="Questions & Answers">
          <Faq
            id="faq-1"
            question="What is the price of OpenSheets?"
            answer="OpenSheets is free, bring your own key and pay directly to the AI provider."
          />
          <Faq
            id="faq-2"
            question="Does OpenSheets store my data?"
            answer="No, OpenSheets does not store your data. It runs entirely in your browser and your data is never sent to our servers."
          />
        </FAQsTwoColumnAccordion>
      </Main>

      <FooterWithNewsletterFormCategoriesAndSocialIcons
        id="footer"
        cta={undefined}
        links={
          <>
            <FooterCategory title="Resources">
              <FooterLink href="/support">Support</FooterLink>
            </FooterCategory>
            <FooterCategory title="Legal">
              <FooterLink href="/privacy">Privacy Policy</FooterLink>
              <FooterLink href="/terms">Terms of Service</FooterLink>
            </FooterCategory>
          </>
        }
        fineprint="© 2025 Oatmeal, Inc."
        socialLinks={
          <>
            <SocialLink href="https://x.com/sionemart" name="X">
              <XIcon />
            </SocialLink>
            <SocialLink
              href="https://github.com/martinsione/opensheets"
              name="GitHub"
            >
              <GitHubIcon />
            </SocialLink>
          </>
        }
      />
    </>
  );
}
