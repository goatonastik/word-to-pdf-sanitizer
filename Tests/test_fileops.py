from pathlib import Path
import sys,tempfile,unittest
from unittest.mock import patch
APP_ROOT=Path(__file__).resolve().parent
if APP_ROOT.name == 'Tests':APP_ROOT=APP_ROOT.parent
if str(APP_ROOT) not in sys.path:sys.path.insert(0,str(APP_ROOT))
import fileops

class Tests(unittest.TestCase):
    def test_working_copy_hashes_source_in_same_pass_and_verifies_destination(self):
        with tempfile.TemporaryDirectory() as directory:
            root=Path(directory);source=root/'source';target=root/'target';source.write_bytes(b'working copy bytes')
            expected=fileops.digest(source)
            self.assertEqual(fileops.copy_hashed(source,target),expected)
            self.assertEqual(fileops.digest(target),expected)

    def test_working_copy_cleanup_on_metadata_failure(self):
        with tempfile.TemporaryDirectory() as directory:
            root=Path(directory);source=root/'source';target=root/'target';source.write_bytes(b'working copy bytes')
            with patch.object(fileops.shutil,'copystat',side_effect=OSError('simulated metadata failure')):
                with self.assertRaises(OSError):fileops.copy_hashed(source,target)
            self.assertTrue(source.exists());self.assertFalse(target.exists())

    def test_interrupted_copy_removes_partial_output(self):
        with tempfile.TemporaryDirectory() as directory:
            root=Path(directory);src=root/'source';dst=root/'output';src.write_bytes(b'original bytes')
            def stop(source,target,*args):target.write(source.read(3));raise OSError('simulated interruption')
            with patch.object(fileops.shutil,'copyfileobj',side_effect=stop):
                with self.assertRaises(OSError):fileops.copy_verified(src,dst)
            self.assertFalse(dst.exists());self.assertEqual(src.read_bytes(),b'original bytes')
    def test_existing_file_is_never_deleted_or_overwritten(self):
        with tempfile.TemporaryDirectory() as directory:
            root=Path(directory);src=root/'source';dst=root/'output';src.write_bytes(b'new');dst.write_bytes(b'existing')
            with self.assertRaises(FileExistsError):fileops.copy_verified(src,dst)
            self.assertEqual(dst.read_bytes(),b'existing')
    def test_checksum_failure_removes_only_new_copy(self):
        with tempfile.TemporaryDirectory() as directory:
            root=Path(directory);src=root/'source';dst=root/'output';src.write_bytes(b'original bytes')
            with self.assertRaises(ValueError):fileops.copy_verified(src,dst,'0'*64)
            self.assertFalse(dst.exists());self.assertTrue(src.exists())

    def fixture(self,root):
        import shutil
        from reportlab.pdfgen.canvas import Canvas
        source=root/'Source.docx';source.write_bytes(b'original Word bytes')
        raw=root/'raw.pdf';clean=root/'clean.pdf';target=root/'Source.pdf';saved=root/'Archive'/'Source.docx'
        canvas=Canvas(str(raw));canvas.drawString(72,720,'Core page');canvas.save()
        shutil.copyfile(raw,str(raw)+'.baseline.pdf')
        return source,raw,clean,target,saved

    def test_failed_pdf_publish_removes_partial_pdf(self):
        import pipeline,convert
        with tempfile.TemporaryDirectory() as directory:
            root=Path(directory);source,raw,clean,target,saved=self.fixture(root)
            def stop(src,dst,*args):dst.write(src.read(80));raise OSError('simulated PDF disk failure')
            with patch.object(fileops.shutil,'copyfileobj',side_effect=stop):
                result=pipeline.verify_and_archive(source,target,saved,raw,clean,convert.digest(source),dict(pipeline.DEFAULTS,archive=str(saved.parent)))
            self.assertEqual(result['status'],'FAILED; SOURCE RETAINED')
            self.assertFalse(target.exists());self.assertFalse(saved.exists())
            self.assertEqual(source.read_bytes(),b'original Word bytes')

    def test_failed_archive_copy_keeps_verified_pdf_and_removes_partial_archive(self):
        import pipeline,convert,shutil
        copy=shutil.copyfileobj
        with tempfile.TemporaryDirectory() as directory:
            root=Path(directory);source,raw,clean,target,saved=self.fixture(root)
            def stop(src,dst,*args):
                if str(dst.name).lower().endswith('.docx'):
                    dst.write(src.read(3));raise OSError('simulated archive disk failure')
                return copy(src,dst,*args)
            with patch.object(fileops.shutil,'copyfileobj',side_effect=stop):
                result=pipeline.verify_and_archive(source,target,saved,raw,clean,convert.digest(source),dict(pipeline.DEFAULTS,archive=str(saved.parent)))
            self.assertEqual(result['status'],'PDF CREATED; SOURCE RETAINED')
            self.assertTrue(target.exists());self.assertFalse(saved.exists())
            self.assertEqual(source.read_bytes(),b'original Word bytes')

    def test_interrupted_restore_cleans_partial_word_file_and_can_retry(self):
        import maintenance,convert
        with tempfile.TemporaryDirectory() as directory:
            root=Path(directory);source=root/'Client';source.mkdir();originals=root/'Originals';archive=originals/'Client';archive.mkdir(parents=True)
            saved=archive/'letter.docx';saved.write_bytes(b'original Word bytes');pdf=source/'letter.pdf';pdf.write_bytes(b'generated PDF')
            maintenance.save_restore_record(source/'letter.docx',pdf,saved,convert.digest(saved),convert.digest(pdf),{'archive':str(archive)})
            plan=maintenance.plan_restore(source,originals)
            def stop(src,dst,*args):dst.write(src.read(3));raise OSError('simulated restore disk failure')
            with patch.object(fileops.shutil,'copyfileobj',side_effect=stop):rows=maintenance.restore(plan)
            self.assertEqual(rows[0][1],'NOT RESTORED');self.assertFalse((source/'letter.docx').exists())
            self.assertEqual(saved.read_bytes(),b'original Word bytes');self.assertEqual(pdf.read_bytes(),b'generated PDF')
            self.assertEqual(maintenance.plan_restore(source,originals)['ready'],1)

if __name__=='__main__':unittest.main()
