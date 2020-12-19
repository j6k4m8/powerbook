# Powerbook

Generate Powerpoint presentations from a Jupyter Notebook.

---

## Background

The built-in Jupyter Notebook Reveal.js slideshow system is nice! I do use it a lot. But it doesn't jive well with collaboration (editors all need to be familiar with the system, and they also need to be able to edit your notebook), and honestly, the industry standard is Powerpoint and that's not going away any time soon. (Beamer sorta exists, but I declare that if you're writing your slideshows in LaTeX you're on your own to get figures into it from Python.)

And hey another thing! My notebook doesn't always match up with the presentation I want to write: coding and storytelling don't always happen in the same order.

## Solution(?)

This (admittedly janky) solution lets you generate a Powerpoint presentation directly from your Jupyter notebooks (or, for that matter, directly from a Python script). When you have a product that you like, such as a figure or an output, add it to your slideshow with `Powerbook#add_image_slide`, for example. Powerbook will automatically figure out how to add matplotlib figures and images from your hard drive.

When you rerun your analysis, your powerpoint file is automatically regenerated with the fresh results.

## Example

```python
from powerbook import Powerbook

P = Powerbook("MySlideshow.pptx")
P.add_title_side("Hello world!", "This is PowerPoint!")
```

For a more in-depth introduction to powerbook from a Jupyter Notebook, see the `examples/` directory of this repository.

## Roadmap

### Slots!

I am dying to figure out an easy way to notate "placeholder" slots in a presentation so that you can send the pptx to your friends to edit and then — rather than overwriting the whole presentation — you can just save changes to specific figures.

Due to the XML schema of pptx, this is super challenging. I wish it weren't so, but this is proving to be much more difficult than I thought.

### More slide templates!

Theoretically you can go off-script and add items to slides anywhere, but it would be nice to have a rich set of slide templates.

### Markdown support!

This is going to be a bear, because markdown will need to be parsed into a set of nested Python objects. Trying to investigate the easiest path forward here, thoughts welcome.

### LaTeX support!

Powerpoint doesn't support this by default even though formulas exist in the app. A temporary shim fix might be to call out to a LaTeX formula image-generator service and inline the image back into the text. But I hate that.

### Syntax-highlighting...

I don't love the idea of lots of code in powerpoint presentations, but if you've gotta add code, it sure better be highlighted.

### Your idea here?
