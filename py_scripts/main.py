import imagehash
import sys
import math
from PIL import Image

if __name__ == '__main__':
    try:
        count = len(sys.argv)
        
        if count != 4:
            raise Exception("Arguments length not correct.")
        
        pic1 = Image.open(sys.argv[1])
        pic2 = Image.open(sys.argv[2])

        width, height = pic1.size
        centerX = width // 2
        centerY = height // 2

        edge = int(sys.argv[3])
    
        half_edge = edge // 2

        left = centerX - half_edge
        if left < 0:
            left = 0

        right = centerX + half_edge
        if right > width:
            right = width

        lower = centerY + half_edge
        if lower > height:
            lower = height

        upper = centerY - half_edge
        if upper < 0:
            upper = 0

        pic1 = pic1.crop((left, upper, right, lower))
        pic2 = pic2.crop((left, upper, right, lower))

        hash1 = imagehash.whash(pic1, hash_size=256)
        hash2 = imagehash.whash(pic2, hash_size=256)

        diff = math.fabs(hash1 - hash2) / hash1.hash.size
        str_diff = f'{diff:.3}'

        sys.stdout.write(str_diff)
        sys.exit(0)

    except Exception as e:
        sys.stderr.write(repr(e))
        sys.exit(-1)
