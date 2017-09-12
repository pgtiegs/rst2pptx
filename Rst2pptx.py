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
    doctree = get_doctree(rststring)
    print(doctree)
    root = ET.fromstring(doctree)
   
    title_slide = prs.slides.add_slide(prs.slide_layouts[0])
    for section in root.findall("section")[:1]:
        print(section.find('title'))
        slide = prs.slides.add_slide(prs.slide_layouts[1])
        title_shape = slide.shapes.title
        title_shape.text = section.find("title").text


    return prs


if __name__ == "__main__":
    print("Here")
