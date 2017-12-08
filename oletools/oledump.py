#!/usr/bin/env python

"""
oledump.py

Dump embedded files

Based on oledump.py (v0.30) by Didier Stevens
(https://blog.didierstevens.com/programs/oledump-py/), part of the
Didier Stevens Suite (https://github.com/DidierStevens/DidierStevensSuite).

Modifications to original include:
- work with streams instead of keeping all data in memory
- dump all files without user-interaction or need re-start with different args
- iterate over zip subfiles and streams as in other oletools files (including
  orphans)
- logging
"""

# === LICENSE =================================================================
#
# Redistribution and use in source and binary forms, with or without
# modification, are permitted provided that the following conditions are met:
#
#  * Redistributions of source code must retain the above copyright notice,
#    this list of conditions and the following disclaimer.
#  * Redistributions in binary form must reproduce the above copyright notice,
#    this list of conditions and the following disclaimer in the documentation
#    and/or other materials provided with the distribution.
#
# THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS"
# AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE
# IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE
# ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE
# LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR
# CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF
# SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS
# INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN
# CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE)
# ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE
# POSSIBILITY OF SUCH DAMAGE.

# -----------------------------------------------------------------------------
# CHANGELOG:
# 2017-12-08 v0.01 CH: - handle all old ole file types (doc, ppt, xls)

# -----------------------------------------------------------------------------
# TODO:
# - dump from ppt
# - create stream for zip sub file
# - check with all file types
# - zip passwords

# -----------------------------------------------------------------------------
# REFERENCES:
# None so far


import sys
import os.path
import logging
import os
from zipfile import is_zipfile, ZipFile
from argparse import ArgumentParser, ArgumentTypeError
import struct

try:
    from oletools.thirdparty import olefile
except ImportError:
    # little hack to allow absolute imports even if oletools is not installed.
    # Copied from olevba.py
    PARENT_DIR = os.path.normpath(os.path.dirname(os.path.dirname(
        os.path.abspath(__file__))))
    if PARENT_DIR not in sys.path:
        sys.path.insert(0, PARENT_DIR)
    del PARENT_DIR
    from oletools.thirdparty import olefile

from ppt_record_parser import is_ppt, PptFile, PptRecordExOleVbaActiveXAtom


# return values from main
RETURN_NO_EXTRACT = 0      # all input files were clean
RETURN_DID_EXTRACT = 1     # did extract files
RETURN_ARGUMENT_ERR = 2    # reserved for parse_args
RETURN_OPEN_FAIL = 3       # failed to open a file
RETURN_STREAM_FAIL = 4     # failed to open an OLE stream

# size of blocks to copy from stream to file
CHUNK_SIZE = 4096    # 4k

# pattern for output files. Will replace: 0 --> count; 1 --> extension
FILE_NAME_PATTERN = 'ole-object-{0:02d}{1}'


def existing_file(filename):
    """ called by argument parser to see whether given file exists """
    if not os.path.isfile(filename):
        raise ArgumentTypeError('{0} is not a file.'.format(filename))
    return filename


def parse_args(cmd_line_args=None):
    """
    parse command line arguments (sys.argv by default)

    Bit of a mess but we want to be compatible with ripOLE, so this can be
    used with amavisd-new
    """
    parser = ArgumentParser(description='Extract files embedded in OLE')
    parser.add_argument('-d', '--target-dir', type=str, default='.',
                        help='Directory to extract files to. File names are '
                             '0.ext, 1.ext ... . Default: current working dir')
    parser.add_argument('-v', '--verbose', action='store_true',
                        help='verbose mode, set logging to DEBUG (else: INFO)')
    parser.add_argument('-i', '--more-input', metavar='FILE',
                        type=existing_file,
                        help='Additional file to parse (same as positional '
                             'arguments)')
    parser.add_argument('input_files', metavar='FILE', nargs='*',
                        type=existing_file,
                        help='Office files to parse (same as -i)')
    args = parser.parse_args(cmd_line_args)

    # combine arguments with -i (compatibility with ripOLE)
    if args.more_input:
        args.input_files += [args.more_input, ]
    if not args.input_files:
        parser.error('No input given (use -i and/or positional argument[s])')

    return args


def find_ole_in_ppt(filename):
    """ find ole streams in ppt """
    for stream in PptFile(filename).iter_streams():
        for record in stream.iter_records():
            if isinstance(record, PptRecordExOleVbaActiveXAtom):
                ole = None
                try:
                    data_start = next(record.iter_uncompressed())
                    if data_start[:len(olefile.MAGIC)] != olefile.MAGIC:
                        continue   # could be an ActiveX control or VBA Storage

                    # otherwise, this should be an OLE object
                    ole = record.get_data_as_olefile()
                    yield ole
                except IOError:
                    logging.warning('Error reading data from {0} stream or '
                                    'interpreting it as OLE object'
                                    .format(stream.name), exc_info=True)
                finally:
                    if ole is not None:
                        ole.close()


def find_ole(filename):
    """ try to open somehow as zip or ole or so; raise exception if fail """
    try:
        if olefile.isOleFile(filename):
            if is_ppt(filename):
                logging.info('is ppt file: ' + filename)
                for ole in find_ole_in_ppt(filename):
                    yield ole
                    ole.close()
            else:
                logging.info('is ole file: ' + filename)
                ole = olefile.OleFileIO(filename)
                yield ole
                ole.close()
        elif is_zipfile(filename):
            logging.info('is zip file: ' + filename)
            zipper = ZipFile(filename, 'r')
            for subfile in zipper.namelist():
                head = b''
                try:
                    with zipper.open(subfile) as file_handle:
                        head = file_handle.read(len(olefile.MAGIC))
                except RuntimeError:
                    logging.error('zip is encrypted: ' + filename)
                    yield None

                if head == olefile.MAGIC:
                    logging.info('  unzipping ole: ' + subfile)
                    with zipper.open(subfile) as file_handle:
                        ole = olefile.OleFileIO(file_handle)
                        yield ole
                        ole.close()
                else:
                    logging.debug('unzip skip: ' + subfile)
        else:
            logging.warning('open failed: ' + filename)
            yield None   # --> leads to non-0 return code
    except Exception:
        logging.error('Caught exception opening {0}'.format(filename),
                      exc_info=True)
        yield None   # --> leads to non-0 return code but try next file first


def ole_iter_streams(ole):
    """
    iterate over streams in open OLE file

    yields 4-tuples (is_orphan, name, size, stream)
    """
    for sid, direntry in enumerate(ole.direntries):
        is_orphan = direntry is None
        if is_orphan:
            # this direntry is not part of the tree --> unused or orphan
            direntry = ole._load_direntry(sid)
        is_stream = direntry.entry_type == olefile.STGTY_STREAM
        logging.debug('direntry {:2d} {}: {}'.format(
            sid, '[orphan]' if is_orphan else direntry.name,
            'is stream of size {}'.format(direntry.size) if is_stream else
            'no stream ({})'.format(direntry.entry_type)))
        if is_stream:
            stream = ole._open(direntry.isectStart, direntry.size)
            yield (is_orphan, None if is_orphan else direntry.name,
                   direntry.size, stream)
            stream.close()


def unpack_stream(format, stream):
    """ struct.unpack with reading and checking of data from stream """
    n_bytes = 0
    for char in format:
        if char in '@=<>!':
            continue   # endianness and stuff
        elif char in 'cbB?':
            n_bytes += 1
        elif char in 'hH':
            n_bytes += 2
        elif char in 'iIlLf':
            n_bytes += 4
        elif char in 'qQd':
            n_bytes += 4
        else:
            raise ValueError('unexpected character "{0}" in format string'
                             .format(char))
    data = stream.read(n_bytes)
    if len(data) != n_bytes:
        raise IOError('reached end of stream after {0} of {1} bytes'
                      .format(len(data), n_bytes))
    return struct.unpack(format, data)

def has_embed_header(stream, stream_size):
    """ check if the header is of expected format

    modeled after oledump.OLE10HeaderPresent(data) but only reads 6 byte
    """

    if stream_size < 6:
        logging.debug('Stream too short ({0})'.format(stream_size))
        return False
    data_size, version = unpack_stream('<LH', stream)
    if data_size == stream_size-4:
        logging.debug('Stream and data size match ({0})'
                      .format(data_size))
    else:
        logging.debug('Stream and data size mismatch {0} != {1}-4'
                      .format(data_size, stream_size))
        return False
    if version != 2:
        logging.debug('Wrong version {0}'.format(version))
        return False
    return True


def read_null_terminated_string(stream):
    """
    Read chars until 0-byte is encountered

    Based on oledump.ReadNullTerminatedString(data), but returns unicode
    """
    chars = []   # array of bytes
    while True:
        data = stream.read(1)
        if len(data) != 1:
            logging.warning('string ended with end-of-stream, not 0 byte')
            break
        char = ord(data[0])
        if char == 0:
            break
        else:
            chars.append(char)
    if sys.version_info.major == 2:
        return u''.join(unichr(char) for char in chars)
    else:
        # have to guess an encoding :-(
        result = None
        chars = bytes(chars)
        for encoding in 'utf16le', 'utf8', 'utf16', 'latin1':
            try:
                result = chars.decode(encoding)
            except UnicodeError:
                pass
            if result is not None:
                return result
        logging.warning('failed to guess encoding for string, falling back to '
                        'ascii with replace')
        return chars.encode('ascii', errors='replace')


def get_embed_info(stream):
    """ get filenames for embedded file from stream

    modeled after oledump.ExtractOle10Native(data) but only reads few bytes

    Returns a size of -1 if size is not contained in stream. That probably
    means that data is not embedded but only linked from file.
    """
    filename = read_null_terminated_string(stream)
    logging.debug('filename: "{0}"'.format(filename))
    pathname = read_null_terminated_string(stream)
    logging.debug('pathname: "{0}"'.format(pathname))
    logging.debug('unused1: {0[0]}, unused2: {0[1]}'
                  .format(unpack_stream('<LL', stream)))
    temppathname = read_null_terminated_string(stream)
    logging.debug('temppathname: "{0}"'.format(temppathname))
    try:
        size_embedded = unpack_stream('<L', stream)[0]
        logging.debug('size embedded: {0}'.format(size_embedded))
    except IOError:
        logging.debug('Data is not embedded (empty size field). '
                      'Probably a link')
        size_embedded = -1

    return (filename, pathname, temppathname), size_embedded


def do_dump(stream, name, embedded_size):
    """ dump data from stream to file with given name, up to embedded_size """
    logging.info('      dumping to ' + name)
    read_count = 0
    with open(name, 'wb') as writer:
        to_read = min(CHUNK_SIZE, embedded_size - read_count)
        while to_read:
            chunk = stream.read(to_read)
            if len(chunk) != to_read:
                logging.warning('Wanted to read {0} but only got {1}'
                                .format(to_read, len(chunk)))
                break
            writer.write(chunk)
            read_count += len(chunk)
            to_read = min(CHUNK_SIZE, embedded_size - read_count)

    if read_count != embedded_size:
        logging.warning('Read count {0} does not match '
                        'expectation {1}'
                        .format(read_count, embedded_size))


def main(cmd_line_args=None):
    """ Main function, called when running file as script

    returns one of the RETURN_* values

    see module doc for more info
    """
    args = parse_args(cmd_line_args)   # does a sys.exit(2) if parsing fails
    if args.verbose:
        logging.basicConfig(level=logging.DEBUG)
    else:
        logging.basicConfig(level=logging.INFO)

    output_count = 0
    return_value = RETURN_NO_EXTRACT

    # loop over file name arguments
    for filename in args.input_files:

        # loop over ole files found within filename
        for ole in find_ole(filename):
            if ole is None:
                return_value = max(return_value, RETURN_OPEN_FAIL)
                continue

            # loop over streams within file
            for is_orphan, name, size, stream in ole_iter_streams(ole):
                logging.info('    Checking stream "{0}"{1}'
                             .format(name, ' (orphan)' if is_orphan else ''))

                # check if this is an embedded file
                if not has_embed_header(stream, size):
                    logging.debug('      not an embedded file - skip')
                    continue

                # get filename options and size of embedded data
                filenames, embedded_size = get_embed_info(stream)
                if embedded_size == -1:
                    continue    # data is not embedded after all

                # make paths compatible with current os
                if os.name in ('posix', 'mac'):  # convert c:\a.ext --> c/a.ext
                    filenames = [fn.replace('\\', '/').replace(':', '/')
                                 .replace('//', '/') for fn in filenames]
                logging.info('      filenames: {0}'.format(filenames))

                # get extension
                extensions = [os.path.splitext(filename)[1].strip()
                              for filename in filenames]
                extensions = [ext for ext in extensions if ext]
                logging.debug('      extensions: {0}'.format(extensions))
                if not extensions:
                    logging.debug('      no extension found, use empty')
                    extension = ''
                elif all(ext == extensions[0] for ext in extensions[1:]):
                    # all extensions are the same
                    logging.debug('      all extension are the same')
                    extension = extensions[0]
                else:
                    logging.debug('      multiple extensions, use first')
                    extension = extensions[0]

                # dump
                name = os.path.join(args.target_dir,
                                    FILE_NAME_PATTERN.format(output_count,
                                                             extension))
                do_dump(stream, name, embedded_size)

                output_count += 1

    if output_count:
        return_value = max(return_value, RETURN_DID_EXTRACT)

    return return_value


if __name__ == '__main__':
    sys.exit(main())
