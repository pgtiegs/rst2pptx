import pptx
import docutils.core
import lxml.etree as ET

def get_doctree(rststring):
    """
    Return the Doctree XML from the RST string
    """
    doctree = docutils.core.publish_string(rststring, writer_name="xml")
    return doctree

def render(rststring):
    prs = pptx.Presentation()
    layouts = dict([(x.name, x) for x in prs.slide_layouts])
    doctree = get_doctree(rststring)
    print(doctree)
    root = ET.fromstring(doctree)
   

    print(root)
    print(layouts)
    title_slide = prs.slides.add_slide(layouts['Title Slide'])

    return prs


if __name__ == "__main__":
    print("Here")
