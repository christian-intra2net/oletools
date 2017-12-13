""" Test oledump basic functionality """

import unittest
from tempfile import mkdtemp
from shutil import rmtree
from os.path import join, isfile
from hashlib import md5
from glob import glob

# Directory with test data, independent of current working directory
from tests.test_utils import DATA_BASE_DIR
from oletools import oledump

DEBUG = True

def calc_md5(filename):
    """ calc md5sum of given file in temp_dir """
    CHUNK_SIZE = 4096
    hasher = md5()
    with open(filename, 'rb') as handle:
        buf = handle.read(CHUNK_SIZE)
        while buf:
            hasher.update(buf)
            buf = handle.read(CHUNK_SIZE)
    return hasher.hexdigest()


class TestOledump(unittest.TestCase):
    """ Tests oledump basic feature """

    def setUp(self):
        """ fixture start: create temp dir """
        self.temp_dir = mkdtemp(prefix='oletools-oledump-')
        self.did_fail = False

    def tearDown(self):
        """ fixture end: remove temp dir """
        if self.did_fail and DEBUG:
            print('leaving temp dir {0} for inspection'.format(self.temp_dir))
        elif self.temp_dir:
            rmtree(self.temp_dir)

    def test_md5(self):
        """ test all files in oledump test dir """
        self.do_test_md5(['-d', self.temp_dir])

    def test_md5_args(self):
        """
        test that oledump can be called with -i and -v

        this is the way that amavisd calls oledump, thinking it is ripOLE
        """
        self.do_test_md5(['-d', self.temp_dir, '-v', '-i'])

    def test_no_output(self):
        """ test that oledump does not find data where it should not """
        args = ['-d', self.temp_dir]
        for sample_name in ('sample_with_lnk_to_calc.doc', ):
            full_name = join(DATA_BASE_DIR, 'oledump', sample_name)
            ret_val = oledump.main(args + [full_name, ])
            if glob(self.temp_dir + 'ole-object-*'):
                self.fail('found embedded data in {0}'.format(sample_name))
            self.assertEqual(ret_val, oledump.RETURN_NO_EXTRACT)

    def do_test_md5(self, args):
        """ helper for test_md5 and test_md5_args """
        # name of sample, extension of embedded file, md5 hash of embedded file
        EXPECTED_RESULTS = (
            ('sample_with_calc_embedded.doc', 'exe', '40e85286357723f326980a3b30f84e4f'),
            ('sample_with_lnk_file.doc', 'lnk', '6aedb1a876d4ad5236f1fbbbeb7274f3'),
            ('sample_with_lnk_file.pps', 'lnk', '6aedb1a876d4ad5236f1fbbbeb7274f3'),
            ('sample_with_lnk_file.ppt', 'lnk', '6aedb1a876d4ad5236f1fbbbeb7274f3'),
        )

        data_dir = join(DATA_BASE_DIR, 'oledump')
        for sample_name, expect_extension, expect_hash in EXPECTED_RESULTS:
            ret_val = oledump.main(args + [join(data_dir, sample_name), ])
            self.assertEqual(ret_val, oledump.RETURN_DID_EXTRACT)
            expect_name = join(self.temp_dir, 'ole-object-00.' + expect_extension)
            if not isfile(expect_name):
                self.did_fail = True
                self.fail('{0} not created from {1}'.format(expect_name,
                                                            sample_name))
                continue
            hash = calc_md5(expect_name)
            if hash != expect_hash:
                self.did_fail = True
                self.fail('Wrong md5 {0} of {1} from {2}'
                          .format(hash, expect_name, sample_name))
                continue


# just in case somebody calls this file as a script
if __name__ == '__main__':
    unittest.main()
