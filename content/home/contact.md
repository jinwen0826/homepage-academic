---
# An instance of the Contact widget.
widget: contact

# This file represents a page section.
headless: true

# Order that this section appears on the page.
weight: 130

title: Contact
subtitle:

content:
  # Automatically link email and phone or display as text?
  autolink: true

  # Email form provider
  form:
    provider: netlify
    formspree:
      id:
    netlify:
      # Enable CAPTCHA challenge to reduce spam?
      captcha: false

  # Contact details (edit or remove options as required)
  email: jinwen@uw.edu
  phone: (+1) 206 200 4742
  address:
    street: 4557 11th Ave NE
    city: Seattle
    region: WA
    postcode: '98105'
    country: United States
    country_code: US
  appointment_url: 'https://calendly.com'
  contact_links:
    - icon: video
      icon_pack: fas
      name: Zoom Me
      link: 'https://zoom.com'

design:
  columns: '2'
---
