from pathlib import Path
import sys,unittest
sys.path.insert(0,str(Path(__file__).resolve().parents[1]))
import convert
import pypdfium2 as pdfium

class Tests(unittest.TestCase):
    def test_equal_pixels_and_single_pixel_difference(self):
        a=pdfium.PdfBitmap.new_native(101,33,pdfium.raw.FPDFBitmap_BGR)
        b=pdfium.PdfBitmap.new_native(101,33,pdfium.raw.FPDFBitmap_BGR)
        try:
            a.fill_rect((255,255,255,255),0,0,101,33);b.fill_rect((255,255,255,255),0,0,101,33)
            self.assertTrue(convert.bitmap_equal(a,b))
            b.fill_rect((0,0,0,255),100,32,1,1)
            self.assertFalse(convert.bitmap_equal(a,b))
        finally:a.close();b.close()

    def test_padding_is_excluded(self):
        a=pdfium.PdfBitmap.new_native(101,3,pdfium.raw.FPDFBitmap_BGR,stride=304)
        b=pdfium.PdfBitmap.new_native(101,3,pdfium.raw.FPDFBitmap_BGR,stride=308)
        try:
            a.fill_rect((255,255,255,255),0,0,101,3);b.fill_rect((255,255,255,255),0,0,101,3)
            self.assertTrue(convert.bitmap_equal(a,b))
        finally:a.close();b.close()

if __name__=='__main__':unittest.main()
