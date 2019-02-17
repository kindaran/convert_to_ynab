import unittest
import convert_to_YNAB_csv as ynab


class Test_Filelist(unittest.TestCase):
    def test_FilelistClean_Valid(self):
        filelist = ynab.getFileList(r"testing\filelist-test1")
        # print(filelist)
        self.assertEqual(len(filelist), 2)
    # end def


# end class

if __name__ == '__main__':
    unittest.main()
