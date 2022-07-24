---
title: Calculate VaR using VBA
— summary: Used KRDs and Convexity
tags:
  - Risk
date: '2022-01-31T00:00:00Z'

# Optional external URL for project (replaces project detail page).
external_link: ''

image:
  caption: Photo by rawpixel on Unsplash
  focal_point: Smart

# links:
# - icon: ""
#    icon_pack: fab
#    name: ""
#   url: ""
url_code: ''
url_pdf: ''
url_slides: ''
url_video: ''



# Slides (optional).
#   Associate this project with Markdown slides.
#   Simply enter your slide deck's filename without extension.
#   E.g. `slides = "example-slides"` references `content/slides/example-slides.md`.
#   Otherwise, set `slides = ""`.
# slides: example
---

I created an Excel tool in the VBA which can calculate VaR for a bond portfolio, by using the historical VaR approach.

In order to calculate the VaR, I calculated the KRDs (Key rate Durations, the partial Durations) and Convexity to be the components of VaR. I calculated each day, each bond VaR in the portfolio, and then added them to get my final VaR for the portfolios.

As for the KRDs, there are 10 KRD points computed for each bond, ranging from 3M-30Years and the KRDs are given in units of KR DV01 at each tenor, which is the change in value for 1 bps decrease of that particular tensor point.

As for the Convexity, I used the OA (option-adjusted) measures convexity for 1 bps decrease in the tenor point.

I calculated the VaR by adding the calculated KRDs and the Convexity in VBA. And then if I put the confidence level, the end date, the interval, and the number of scenarios in Excel, I can get the according VaR by using my wrote VBA functions.



