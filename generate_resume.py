import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn, nsdecls

def create_resume():
    doc = docx.Document()
    
    # Set page margins to 0.5 in for clean compact modern layout
    for section in doc.sections:
        section.top_margin = Inches(0.5)
        section.bottom_margin = Inches(0.5)
        section.left_margin = Inches(0.55)
        section.right_margin = Inches(0.55)
        
    # Color palette
    COLOR_PRIMARY = RGBColor(139, 29, 52)     # Burgundy #8B1D34
    COLOR_DARK = RGBColor(26, 26, 26)         # MCE Black #1A1A1A
    COLOR_MUTED = RGBColor(85, 85, 85)        # Muted Charcoal #555555
    COLOR_ACCENT = RGBColor(85, 107, 47)      # Olive #556B2F
    COLOR_BLUE = RGBColor(30, 90, 160)        # Link blue
    
    def set_font(run, name='Calibri', size=10.5, color=COLOR_DARK, bold=False, italic=False):
        run.font.name = name
        run.font.size = Pt(size)
        run.font.color.rgb = color
        run.bold = bold
        run.italic = italic

    def add_heading_1(text):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(9)
        p.paragraph_format.space_after = Pt(3)
        p.paragraph_format.keep_with_next = True
        run = p.add_run(text.upper())
        set_font(run, name='Calibri', size=11, color=COLOR_PRIMARY, bold=True)
        
        # Add bottom border to paragraph
        pPr = p._p.get_or_add_pPr()
        pBdr = parse_xml(r'<w:pBdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
                         r'<w:bottom w:val="single" w:sz="8" w:space="2" w:color="8B1D34"/>'
                         r'</w:pBdr>')
        pPr.append(pBdr)
        return p

    def add_bullet(text_bold, text_normal, space_after=2):
        p = doc.add_paragraph(style='List Bullet')
        p.paragraph_format.space_before = Pt(1)
        p.paragraph_format.space_after = Pt(space_after)
        p.paragraph_format.left_indent = Inches(0.25)
        if text_bold:
            r1 = p.add_run(text_bold)
            set_font(r1, name='Calibri', size=9.5, color=COLOR_DARK, bold=True)
        if text_normal:
            r2 = p.add_run(text_normal)
            set_font(r2, name='Calibri', size=9.5, color=COLOR_DARK, bold=False)
        return p

    # ==================== HEADER ====================
    p_name = doc.add_paragraph()
    p_name.paragraph_format.space_before = Pt(0)
    p_name.paragraph_format.space_after = Pt(1)
    p_name.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_name = p_name.add_run("BHUMIKA PATEL")
    set_font(r_name, name='Calibri', size=20, color=COLOR_PRIMARY, bold=True)

    p_title = doc.add_paragraph()
    p_title.paragraph_format.space_before = Pt(0)
    p_title.paragraph_format.space_after = Pt(4)
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_title = p_title.add_run("Senior Software Engineer  |  Full-Stack & eCommerce Solutions")
    set_font(r_title, name='Calibri', size=11, color=COLOR_DARK, bold=True)

    p_contact = doc.add_paragraph()
    p_contact.paragraph_format.space_before = Pt(0)
    p_contact.paragraph_format.space_after = Pt(6)
    p_contact.alignment = WD_ALIGN_PARAGRAPH.CENTER
    contact_parts = [
        ("📞 +91-8780432116", False),
        ("  |  ", False),
        ("✉ bhumikapatel8511@gmail.com", False),
        ("  |  ", False),
        ("🌐 bhumika081.github.io/portfolio", False),
        ("  |  ", False),
        ("💻 github.com/bhumika081", False),
        ("  |  ", False),
        ("📍 Ahmedabad, Gujarat, India", False)
    ]
    for text, bold in contact_parts:
        r = p_contact.add_run(text)
        set_font(r, name='Calibri', size=9, color=COLOR_MUTED, bold=bold)

    # ==================== PROFESSIONAL SUMMARY ====================
    add_heading_1("Professional Summary")
    p_sum = doc.add_paragraph()
    p_sum.paragraph_format.space_before = Pt(2)
    p_sum.paragraph_format.space_after = Pt(4)
    r_sum = p_sum.add_run(
        "Results-driven Senior Software Engineer with 4+ years of professional expertise engineering scalable full-stack web applications, high-conversion eCommerce platforms, and custom SaaS tools across service and product environments. Extensive mastery spanning modern frontend (React, Next.js, Angular, TypeScript), robust backend architectures (PHP, CodeIgniter 4, Node.js), enterprise eCommerce systems (Shopify App Bridge & Polaris, Magento 2, WooCommerce), and workflow automations (n8n, GoHighLevel). Shipped 19+ commercial web platforms across 6 international markets (USA, UK, Australia, Germany, Ireland, and India) with strict adherence to WCAG 2.1 accessibility, clean modular architecture, and sub-second performance benchmarks."
    )
    set_font(r_sum, name='Calibri', size=9.5, color=COLOR_DARK)

    # ==================== TECHNICAL SKILLS ====================
    add_heading_1("Core Technical Skills")
    skills_data = [
        ("Frontend Technologies: ", "React.js, Next.js (SSR / SSG / App Router), Angular 12+, TypeScript, JavaScript (ES6+), HTML5, CSS3, Tailwind CSS, Bootstrap 5, SASS/SCSS, AJAX, jQuery, Responsive Web Design."),
        ("Backend & Architectures: ", "PHP (7.4 – 8.3), CodeIgniter 4, Node.js, Express, RESTful API Engineering, MVC Architecture, Role-Based Access Control (RBAC), Object-Oriented Programming (OOP)."),
        ("eCommerce & CMS Platforms: ", "Shopify App Development (Shopify App Bridge, Polaris UI, GraphQL Admin API, Liquid), Magento 2 (CE/EE, Custom Modules, Catalog Sync), WooCommerce, WordPress (Custom Theme & Plugin Development, ACF Pro, Gutenberg)."),
        ("Databases & Cloud Infrastructure: ", "MySQL, PostgreSQL, Database Schema Design, Query Optimization, Indexing, Git, GitHub, Vercel, cPanel, Apache, Nginx, Linux Server Administration."),
        ("Integrations & Automations: ", "Stripe API & Subscriptions, PayPal, OpenAI API, Google Maps Platform, Twilio, Bookly API, n8n Workflow Automations, GoHighLevel (GHL) CRM, Klaviyo."),
        ("Tools, Testing & Methodologies: ", "Agile/Scrum, Full SDLC Management, Postman, VS Code, Asana, Jira, Slack, Cross-Browser Compatibility Testing, Lighthouse Auditing & Speed Optimization, WCAG 2.1 Accessibility.")
    ]
    for label, val in skills_data:
        add_bullet(label, val, space_after=1.5)

    # ==================== WORK EXPERIENCE ====================
    add_heading_1("Professional Work Experience")

    # Job 1
    p_job1 = doc.add_paragraph()
    p_job1.paragraph_format.space_before = Pt(3)
    p_job1.paragraph_format.space_after = Pt(1)
    p_job1.paragraph_format.keep_with_next = True
    r_role1 = p_job1.add_run("Senior Software Engineer  |  Webile Technologies Pvt. Ltd")
    set_font(r_role1, name='Calibri', size=10, color=COLOR_DARK, bold=True)
    r_date1 = p_job1.add_run("  —  Ahmedabad, India")
    set_font(r_date1, name='Calibri', size=9.5, color=COLOR_MUTED, italic=True)
    p_job1_right = p_job1.add_run("  [July 2022 – Present]")
    set_font(p_job1_right, name='Calibri', size=9.5, color=COLOR_PRIMARY, bold=True)

    add_bullet("Full-Stack Leadership: ", "Architect, develop, and maintain enterprise web applications, high-traffic eCommerce storefronts, and custom portal solutions for global clients across US, UK, Australia, Germany, and Ireland.")
    add_bullet("Shopify App Ecosystem: ", "Engineered and launched ShopPilot AI on the official Shopify App Store using Shopify App Bridge, Polaris UI, Node.js, and OpenAI API, automating merchant catalog and inventory management.")
    add_bullet("Enterprise Magento & eCommerce: ", "Led custom module engineering and payment gateway integrations for enterprise Magento 2 systems (Tesselaar Flowers Australia, DigiHorizon Store) and WooCommerce platforms (Maguire Supplies Ireland, La Fleur De Luxe USA).")
    add_bullet("Frontend Modernization: ", "Constructed high-speed single-page applications and calculators in Angular 12+ and React.js for major financial institutions (Patelco Credit Union) and healthcare/wellness providers (Wellness & Health 4 All), achieving 98+ Google Lighthouse scores.")
    add_bullet("Custom API & Workflow Automations: ", "Designed resilient webhook integrations connecting eCommerce stores with n8n workflow engines, GoHighLevel CRM, Stripe subscription billing, and third-party fulfillment services.")
    add_bullet("Code Quality & Mentorship: ", "Enforce Git branching strategies, perform peer code reviews, optimize MySQL schemas and query execution plans, and mentor junior developers in modern full-stack practices.")

    # Job 2
    p_job2 = doc.add_paragraph()
    p_job2.paragraph_format.space_before = Pt(4)
    p_job2.paragraph_format.space_after = Pt(1)
    p_job2.paragraph_format.keep_with_next = True
    r_role2 = p_job2.add_run("PHP Developer  |  LIOM IT Hub Pvt. Ltd")
    set_font(r_role2, name='Calibri', size=10, color=COLOR_DARK, bold=True)
    r_date2 = p_job2.add_run("  —  Ahmedabad, India")
    set_font(r_date2, name='Calibri', size=9.5, color=COLOR_MUTED, italic=True)
    p_job2_right = p_job2.add_run("  [May 2021 – December 2021]")
    set_font(p_job2_right, name='Calibri', size=9.5, color=COLOR_PRIMARY, bold=True)

    add_bullet("Backend & Database Engineering: ", "Developed backend logic, custom controllers, and relational MySQL schemas for web portals using PHP, MVC frameworks, and WordPress.")
    add_bullet("Security & Access Control: ", "Engineered secure role-based access control (RBAC), session management, custom AJAX data filtering, and administrative management dashboards.")
    add_bullet("Theme & Plugin Development: ", "Engineered custom WordPress themes from scratch and developed bespoke plugins utilizing Advanced Custom Fields (ACF) to support dynamic client requirements.")

    # ==================== ALL 20 PRODUCTION PROJECTS ====================
    add_heading_1("Comprehensive Production Projects Portfolio (20 Live Applications)")

    projects = [
        # 1
        ("ShopPilot AI  (Official Shopify App Store)", "https://apps.shopify.com/shoppilot-ai-1",
         "Shopify App Bridge, Polaris UI, Node.js, React, OpenAI API, GraphQL Admin API",
         "Published official Shopify application automating product copy, SEO tagging, and smart inventory suggestions. Integrated App Bridge embedded authentication, webhook subscriptions, and secure billing via Shopify GraphQL Billing API."),
        # 2
        ("Tesselaar Flowers Australia", "https://www.tesselaar.net.au/",
         "Magento 2 Enterprise eCommerce, Custom PHP Modules, B2B/B2C Checkout, ERP Sync",
         "Australia's foremost flower nursery portal. Architected custom catalog sync for thousands of seasonal bulb SKUs, pre-order reservation logic, complex Australian shipping matrix calculation, and automated inventory reconciliation."),
        # 3
        ("DigiHorizon Store", "https://store.digihorizon.in/",
         "Magento 2 eCommerce, MySQL, Payment Gateways, High-Performance Caching",
         "Digital goods and electronics commercial marketplace. Built multi-category product catalog, secure Indian and international payment gateways, warranty registration system, and Redis caching for sub-second page loads."),
        # 4
        ("Patelco Credit Union – Loan Calculator", "https://www.patelco.org/",
         "Angular 12+, TypeScript, Reactive Forms, Financial Algorithms, CSS3",
         "Interactive financial loan calculation suite for 300,000+ member California credit union. Implemented dynamic monthly payment sliders, amortization projection graphs, and seamless lead handoff to loan officers."),
        # 5
        ("Patelco Credit Union – Enterprise Web Portal", "https://www.patelco.org/",
         "WordPress Enterprise, Advanced Custom Fields (ACF Pro), Bookly API, WCAG 2.1",
         "Large-scale enterprise banking web portal. Built bespoke custom Gutenberg blocks, branch appointment scheduling integration, ADA/WCAG 2.1 accessibility compliance, and high-availability WordPress architecture."),
        # 6
        ("Ethical Hire Australia", "https://ethicalhire.com.au/",
         "CodeIgniter 4 (PHP), MySQL, Stripe API Subscriptions, Google Maps API",
         "Comprehensive Australian recruitment marketplace. Developed multi-tenant RBAC (Candidates, Employers, Admins), recurring monthly Stripe subscription billing, geolocation radius search, and applicant tracking workflows."),
        # 7
        ("TRLIT Solutions", "https://trlit.com/",
         "Next.js, React, Node.js API, Server-Side Rendering (SSR), Tailwind CSS",
         "Enterprise cloud IT solutions platform. Built scalable Next.js architecture with SSR for rapid content delivery, dynamic capability showcase, and automated lead capture pipelines."),
        # 8
        ("The Spa of Charleston", "https://www.thespaofcharleston.com/",
         "WordPress, Custom PHP, ACF Pro, Appointment Booking Engine, eCommerce",
         "Historic South Carolina luxury wellness portal. Engineered online treatment discovery menu, custom appointment scheduling workflow, instant gift certificate generation, and secure checkout."),
        # 9
        ("Ausstellungen – Gitta Marquardt-Baladurage", "https://gitta-marquardt-baladurage.de/ausstellungen/",
         "WordPress, Custom PHP Theme, Fine Arts Archive, Responsive UI",
         "German contemporary fine arts and retrospective exhibition platform. Engineered curated art gallery archives, chronological exhibition timelines, and bilingual high-resolution responsive layout."),
        # 10
        ("My Jewellery Story", "https://www.myjewellerystory.com.au/",
         "Shopify Liquid, Custom JavaScript/CSS, High-Res Zoom, Klaviyo",
         "Australian luxury jewelry eCommerce. Engineered high-resolution gemstone magnification viewer, real-time metal/stone variant swatches, and high-converting cart drawer upsells."),
        # 11
        ("Fashion Jewellery Hub", "https://fashionjewelleryhub.myshopify.com/",
         "Shopify Theme Architecture, AJAX Filtering, Quick-Buy Drawer",
         "High-velocity lifestyle fashion jewelry brand. Implemented instant AJAX facet search by metal, price, and occasion, optimized 1-click mobile checkout, and live Instagram feed integration."),
        # 12
        ("Halal 4 All Shop", "https://halal4allshop.com/",
         "Shopify Liquid, Storefront API, Local Delivery Slots, Payment Gateways",
         "International Halal food and grocery eCommerce. Built multi-region currency conversions, local zip-code delivery scheduler, cold-storage shipping rules, and payment gateways."),
        # 13
        ("Maguire Supplies Ireland", "https://maguiresupplies.com/",
         "WooCommerce, Custom PHP, B2B Wholesale Pricing, EU VAT Engine",
         "Irish industrial supplies and hardware portal. Developed B2B wholesale portal with approved trade logins, tiered bulk volume discounts, automated EU VAT validation, and PDF invoice generation."),
        # 14
        ("Retail Matchmaker", "https://retailmatchmaker.com/",
         "Full-Stack Web App, PHP, MySQL, Brand-Retailer Match Engine",
         "B2B matchmaking platform connecting emerging consumer brands with nationwide department stores and buyers. Features custom application vetting forms, direct messaging, and buyer discovery filters."),
        # 15
        ("SparknFly Canada", "https://sparknfly.ca/",
         "WordPress, Custom PHP Admin Dashboard, Role-Based Access Control",
         "Licensed Canadian dispensary platform (21+). Designed custom PHP backend dashboard for dispensary operators, age verification gateway, real-time inventory counts, and AJAX product filtering."),
        # 16
        ("WebRX Pharma UK", "https://webrx.co.uk/",
         "WordPress, HIPAA/GDPR Compliance, Prescription Intake Engine",
         "Regulated UK digital pharmacy. Developed secure multi-step clinical health questionnaire, doctor approval authorization queue, electronic prescription dispensing states, and discreet packaging options."),
        # 17
        ("La Fleur De Luxe USA", "https://lafleurdeluxe.com/",
         "WordPress / WooCommerce, Custom Box Configurator, Stripe Checkout",
         "Luxury preserved rose florist. Engineered bespoke arrangement visualizer (box shape, lid finish, floral arrangement, custom embossed greeting cards), and streamlined checkout."),
        # 18
        ("Wellness & Health 4 All", "https://wellnessandhealth4all.com/",
         "React.js, Tailwind CSS, REST APIs, Context API State Management",
         "Mindfulness retreat and wellness event booking platform. Built retreat discovery explorer with multi-criteria filtering, instructor bio modal showcases, and dynamic registration workflows."),
        # 19
        ("Javia International", "https://javiainternational.com/",
         "WordPress, Custom PHP Theme, Multilingual Architecture, SEO Optimization",
         "Global corporate import/export trade portal. Engineered multi-category industrial product directory, multilingual localization, RFP quotation request workflow, and international SEO foundation."),
        # 20
        ("NVD Technology", "https://nvdtechnology.com/",
         "WordPress, Custom CSS, Interactive Service Matrices, Lead Funnel",
         "Corporate IT consulting and technology advisory website. Features interactive service solution matrices, customer case study spotlights, and integrated CRM inquiry pipelines.")
    ]

    for title, url, tech, desc in projects:
        p_proj = doc.add_paragraph()
        p_proj.paragraph_format.space_before = Pt(3)
        p_proj.paragraph_format.space_after = Pt(1)
        p_proj.paragraph_format.keep_with_next = True
        
        r_t = p_proj.add_run(f"•  {title}")
        set_font(r_t, name='Calibri', size=9.5, color=COLOR_DARK, bold=True)
        
        r_url = p_proj.add_run(f"  [{url}]")
        set_font(r_url, name='Calibri', size=8.5, color=COLOR_BLUE, italic=True)
        
        p_desc = doc.add_paragraph()
        p_desc.paragraph_format.space_before = Pt(0)
        p_desc.paragraph_format.space_after = Pt(3)
        p_desc.paragraph_format.left_indent = Inches(0.2)
        
        r_tech = p_desc.add_run("Tech Stack: ")
        set_font(r_tech, name='Calibri', size=8.5, color=COLOR_PRIMARY, bold=True)
        r_tech_val = p_desc.add_run(f"{tech}\n")
        set_font(r_tech_val, name='Calibri', size=8.5, color=COLOR_MUTED, italic=True)
        
        r_d = p_desc.add_run(desc)
        set_font(r_d, name='Calibri', size=9, color=COLOR_DARK)

    # ==================== EDUCATION ====================
    add_heading_1("Education & Academic Background")
    p_edu = doc.add_paragraph()
    p_edu.paragraph_format.space_before = Pt(3)
    p_edu.paragraph_format.space_after = Pt(2)
    
    r_deg = p_edu.add_run("Bachelor of Engineering (B.E.) in Computer Engineering")
    set_font(r_deg, name='Calibri', size=10, color=COLOR_DARK, bold=True)
    r_uni = p_edu.add_run("\nApollo Institute of Engineering and Technology  |  Gujarat Technological University (GTU)")
    set_font(r_uni, name='Calibri', size=9.5, color=COLOR_MUTED)
    r_year = p_edu.add_run("  —  Graduated: 2020")
    set_font(r_year, name='Calibri', size=9.5, color=COLOR_PRIMARY, bold=True)

    # Save document
    output_path = r'd:\Bhumika Patel\Bhumika_Patel_Senior_Software_Engineer_Resume.docx'
    doc.save(output_path)
    print(f"Successfully generated resume: {output_path}")

if __name__ == '__main__':
    create_resume()
