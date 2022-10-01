import imageio
import glob

def cropDialogue(cropimg, index):
    if cropimg[8, 237, 0] >= 88 and cropimg[8, 237, 0] <= 93:
        imgcroped = cropimg[3:129, 236:565]
        imageio.imwrite('F:/Dia2/dialogue/%04d.png' % index, imgcroped)

def cropWindows(filepath, width = 1, height=26, wsize=800, hsize=600):
    imglist = glob.glob(filepath + '/*.png')
    for i, name in enumerate(imglist):
        img = imageio.imread(imglist[i])
        cropimg = img[height: height + hsize, width: width + wsize]
        imageio.imwrite('F:/Dia2/%04d.png' %i, cropimg)
        cropDialogue(cropimg, i)

if __name__ == "__main__":
    cropWindows('F:/Dia2')