import unittest
import rst2pptx
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
        prs = rst2pptx.render(sample)

        self.assertIsInstance(prs, pptx.presentation.Presentation )

        self.assertEqual(prs.slides[0].slide_layout.name, "Title Slide")
        
       
        print([x.name for x in prs.slide_layouts])
        prs.save("test.pptx")

if __name__ == '__main__':
    unittest.main()
