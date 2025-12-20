import { ButtonLink, PlainButtonLink } from "@/components/elements/button";
import { Container } from "@/components/elements/container";
import { Document } from "@/components/elements/document";
import { Main } from "@/components/elements/main";
import { GitHubIcon } from "@/components/icons/social/github-icon";
import { XIcon } from "@/components/icons/social/x-icon";
import {
  FooterCategory,
  FooterLink,
  FooterWithNewsletterFormCategoriesAndSocialIcons,
  NewsletterForm,
  SocialLink,
} from "@/components/sections/footer-with-newsletter-form-categories-and-social-icons";
import {
  NavbarLink,
  NavbarLogo,
  NavbarWithLinksActionsAndCenteredLogo,
} from "@/components/sections/navbar-with-links-actions-and-centered-logo";

export default function LegalLayout({
  children,
}: {
  children: React.ReactNode;
}) {
  return (
    <>
      <NavbarWithLinksActionsAndCenteredLogo
        id="navbar"
        links={undefined}
        logo={
          <NavbarLogo href="/">
            <p className="font-serif text-2xl">OpenSheets</p>
          </NavbarLogo>
        }
        actions={
          <>
            <PlainButtonLink href="#" className="max-sm:hidden">
              Log in
            </PlainButtonLink>
            <ButtonLink href="#">Get started</ButtonLink>
          </>
        }
      />

      <Main>
        <section className="py-16">
          <Container>
            <Document className="mx-auto max-w-2xl">{children}</Document>
          </Container>
        </section>
      </Main>

      <FooterWithNewsletterFormCategoriesAndSocialIcons
        id="footer"
        cta={
          <NewsletterForm
            headline="Stay in the loop"
            subheadline={
              <p>
                Get product updates and tips delivered straight to your inbox.
              </p>
            }
            action="#"
          />
        }
        links={
          <>
            <FooterCategory title="Product">
              <FooterLink href="#">Features</FooterLink>
              <FooterLink href="#">Pricing</FooterLink>
              <FooterLink href="#">Integrations</FooterLink>
            </FooterCategory>
            <FooterCategory title="Company">
              <FooterLink href="#">About</FooterLink>
              <FooterLink href="#">Blog</FooterLink>
            </FooterCategory>
            <FooterCategory title="Resources">
              <FooterLink href="https://github.com/martinsione/opensheets">
                GitHub
              </FooterLink>
              <FooterLink href="/support">Support</FooterLink>
            </FooterCategory>
            <FooterCategory title="Legal">
              <FooterLink href="/privacy">Privacy Policy</FooterLink>
              <FooterLink href="/terms">Terms of Service</FooterLink>
            </FooterCategory>
          </>
        }
        fineprint="© 2025 OpenSheets"
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
