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
"""
        prs = rst2pptx.render(sample)

        self.assertIsInstance(prs, pptx.presentation.Presentation )

        self.assertEqual(prs.slides[0].slide_layout.name, "Title Slide")
        
       
        print([x.name for x in prs.slide_layouts])


if __name__ == '__main__':
    unittest.main()
