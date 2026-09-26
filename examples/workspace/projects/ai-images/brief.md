# Brief: "Our new waterfront studio"

**Audience.** Anyone trying pptx-gen's optional image generation for the first
time.

**Objective.** Two slides that carry generated photography: a cover built from
the `title-cover` template with a picture beside the title, and a custom slide
that pairs a short message with a second picture. Together they show both ways a
generated image is placed, and what the deck looks like before the images exist.

**How to use it.**

1. `pptx-gen images --project examples/workspace/projects/ai-images` generates
   the images below into `inputs/` (it needs `OPENAI_API_KEY`).
2. `pptx-gen build --script examples/workspace/projects/ai-images/build.ts`
   places them. Without the files, each slide shows a captioned grey
   placeholder instead.
3. `output/report.md` shows every variant side by side. Change `pick` below and
   rebuild to swap one in; that never regenerates anything.

## Visuals

- Slide 1 — The harbour at first light, beside the title → generated image (`harbour-dawn`)
- Slide 2 — The team at work in the new space → generated image (`studio-interior`)

## Images

```yaml
- id: harbour-dawn
  slide: 1
  description: A quiet harbour at first light, seen from the waterfront; moored sailing boats on calm water and a low sun rising over the far shore
  variants: 3
  pick: 1
  aspect: portrait
- id: studio-interior
  slide: 2
  description: A bright open-plan design studio in a converted harbour warehouse, people sketching together at a long wooden table, tall windows looking out onto the water
  aspect: square
```
