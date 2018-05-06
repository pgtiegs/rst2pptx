import logging
import pptx
import docutils.core
import lxml.etree as ET

logging.basicConfig(level=logging.DEBUG)

def get_doctree(rststring):
    """
    Return the Doctree XML from the RST string
    """
    doctree = docutils.core.publish_string(rststring, writer_name="xml")
    return doctree

def render(rststring):
    prs = pptx.Presentation()
    doctree = get_doctree(rststring)
    logging.debug(doctree)
    root = ET.fromstring(doctree)
   
    title_slide = prs.slides.add_slide(prs.slide_layouts[0])
    title_shape = title_slide.shapes.title
    title_shape.text = root.find('title').text
    for section in root.findall("section"):
        logging.debug(section.find('title').text)
        slide = prs.slides.add_slide(prs.slide_layouts[1])
        for child in section:
            if child.tag == 'bullet_list':
                logging.debug("Bullet List Handler")
                logging.debug([x.name for x in slide.shapes])
            elif child.tag == 'title':
                logging.debug("Title Handler")
                title_shape = slide.shapes.title
                title_shape.text = section.find("title").text
            elif child.tag == 'section':
                logging.debug("Sub-Section Handler")
            else:
                logging.debug("New tag: {}".format(child.tag))


    return prs


if __name__ == "__main__":
    fd = open("test.rst", "r")
    render(fd.read()).save("test.pptx")
    fd.close()
