import olefile
import shutil
import os
import zipfile
import tempfile
from glob import  iglob
from pathlib import Path

def stomp_vba(original_file, stomped_file):
    
    if olefile.isOleFile(original_file):
        # Make copy of file to modify.
        shutil.copyfile(original_file, stomped_file)
        stomp_it(stomped_file)
    elif zipfile.is_zipfile(original_file):
        #unzip to temporary location
        tmpdir = tempfile.TemporaryDirectory(prefix="stomp_")
        with zipfile.ZipFile(original_file) as zf:
            zf.extractall(tmpdir.name)
            #iterate through files, call stomp on any ole file
            file_list = [f for f in iglob(tmpdir.name + '/**/*', recursive=True) if os.path.isfile(f)]
            for f in file_list:
                if olefile.isOleFile(f):
                    stomp_it(f)
            #write the new zip file
            os.chdir(Path(stomped_file).resolve().parent)
            shutil.make_archive(stomped_file, 'zip', tmpdir.name)
            if os.path.exists(stomped_file): os.remove(stomped_file)
            os.rename(stomped_file + '.zip',stomped_file)

# This method originally from Kirk Sayre (@bigmacjpg)             
def stomp_it(stomped_file):    
    # Open file to mangle the VBA streams.
    with olefile.OleFileIO(stomped_file, write_mode=True) as ole:
    
        # Mangle the macro VBA streams.
        for stream_name in ole.listdir():
        
            # Got macros?
            data = ole.openstream(stream_name).read()
            
            marker = b"Attrib"
            if (marker not in data):
                continue
           
            # Find where to write.
            start = data.rindex(marker)

            # Stomp the rest of the data.
            new_data = data[:start]
            for i in range(start, len(data)):
                # Stomp with random bytes.
                new_data += os.urandom(1)
            
            # Write out the garbage data.
            ole.write_stream(stream_name, new_data)


    