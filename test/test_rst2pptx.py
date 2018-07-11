import unittest
import docutils.core
import docutils.io

import pptx
import rst2pptx

class Test_rst2pptx(unittest.TestCase):
    def test_basic_read_write(self):
        sample = """
=====
Title
=====

My Title Slide
==============
* This is slide

Second Slide
============

* Hello

sub-section
-----------
My Subsection

"""
        writer = rst2pptx.PowerPointWriter()
        writer.document = docutils.core.publish_doctree(sample)
        writer.presentation = pptx.Presentation()
        writer.translate()

        self.assertIsInstance(writer.presentation, pptx.presentation.Presentation )

        self.assertEqual(writer.presentation.slides[0].slide_layout.name, "Title Slide")
        
       
        print([x.name for x in writer.presentation.slide_layouts])
        writer.presentation.save("test.pptx")

if __name__ == '__main__':
    unittest.main()
