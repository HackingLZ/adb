from argparse import ArgumentParser
from argparse import RawDescriptionHelpFormatter
import sys
import build
import os
from pathlib import Path
from stomp_vba import stomp_vba
from glob import  iglob
from stomp_vba import stomp_vba

def get_args(arg_list):
    del arg_list[0]
    try:
        parser = ArgumentParser(description='program description', formatter_class=RawDescriptionHelpFormatter, prog='program_name')
        parser.add_argument('-a', '--adversary', dest="adversary", default=None, help="-a --adversary {adversary name} (use -l to list)", required=False)
        parser.add_argument('-f', '--filetype', dest="filetype", default="doc",
                            help="-f --filetype doc | docm | flatxml", required=False)
        parser.add_argument('-e', '--extension', dest="extension", default="doc", required=False,
                            help="-e --extension doc | docm | rtf")
        parser.add_argument('-c', '--count', dest="count", default=1, help="-c --count {# of docs to create}", required=False)
        parser.add_argument('-l', '--listadversaries', dest="listadversaries", action='store_true', help="-l --listadversaries : list available adversaries and exits")
        parser.add_argument('-o', '--outdir', dest="outdir", default="Default", help="-o --outdir {path\\to\\outdir}")
        parser.add_argument('-d', '--debug', dest="debug", default=False, action='store_true', help="-d --debug : print debug statements and playbook for each document")
        parser.add_argument('-v', '--vbastomp', dest="vbastomp", default=None, help="-v --vba-stomp : a file or directory of files in which to stomp the vba source")
        args = parser.parse_args(arg_list)
    except KeyboardInterrupt:
        sys.exit(0)
    except Exception as e:
        print('[-] fatal error parsing arguments, error=' + repr(e) + ". for help please user --help")
        raise
    return args


args = get_args(sys.argv)

if args.listadversaries == True:
    adversary_list = os.listdir(os.path.join(Path(__file__).parent, "adversary"))
    for item in adversary_list:
        if os.path.isdir(os.path.join(Path(__file__).parent, "adversary", item)) and "__" not in item:
            print(item)
    quit()
       
#check output dir
if os.path.isdir(args.outdir) == False:
    print("\n[!] outdir is either not defined or is not a valid folder path\n")
    raise ValueError

if args.adversary:

    build.build_files(
        adversary=args.adversary,
        out_dir=args.outdir,
        count=int(args.count),
        filetype=args.filetype,
        extension=args.extension,
        debug=args.debug
    )
elif args.vbastomp:
    files = []
    if os.path.isdir(args.vbastomp):
        files = [f for f in iglob(args.vbastomp + '/*', recursive=False) if os.path.isfile(f)]
    elif os.path.isfile(args.vbastomp) == False:
        print("\n[!] specified file does not exist\n")
        raise ValueError
    else:
        files.append(args.vbastomp)
    
    for f in files:
        new_filename = os.path.join(Path(args.outdir),Path(f).resolve().stem + "_stomped" + Path(f).resolve().suffix)
        stomp_vba(f, new_filename)
        print("DONE! see " + new_filename)
    
