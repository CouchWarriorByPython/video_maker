import os
import random
import openpyxl
import numpy as np
from moviepy.editor import AudioFileClip, ImageClip, TextClip, CompositeVideoClip, concatenate_videoclips
from moviepy.video.fx.all import resize
from moviepy.video.tools.segmenting import findObjects

workbook = openpyxl.load_workbook('1.xlsx')
sheet = workbook.active
words = []
for cell in sheet[2]:
    words.append(cell.value)

fixed_image_duration = 4
screensize = (1080, 1920)
rotMatrix = lambda a: np.array([[np.cos(a), np.sin(a)], [-np.sin(a), np.cos(a)]])
image_folder = "images"
image_files = [os.path.join(image_folder, f) for f in os.listdir(image_folder) if f.endswith(".jpeg")]
image_files.sort()

def vortex(screenpos, i, nletters):
    scale = 2
    d = lambda t: 1.0 / (0.1 + (t * scale)**8)
    a = i * np.pi / nletters
    v = rotMatrix(a).dot([-1, 0])
    if i % 2:
        v[1] = -v[1]
    return lambda t: screenpos + 800 * d(t) * rotMatrix(0.5 * d(t) * a).dot(v)

def cascade(screenpos, i, nletters):
    scale = 2
    v = np.array([0, -1000])
    d = lambda t: 1 if (t * scale) < 0 else abs(np.sinc(t * scale) / (1 + (t * scale)**2))
    return lambda t: screenpos + v * d(t - 0.05 * i)

def arrive(screenpos, i, nletters):
    scale = 2
    v = np.array([-1, 0])
    d = lambda t: max(0, 3 - 3 * (t * scale))
    return lambda t: screenpos - 800 * v * d(t - 0.1 * i)

def vortexout(screenpos, i, nletters):
    scale = 2
    d = lambda t: max(0, t * scale)
    a = i * np.pi / nletters
    v = rotMatrix(a).dot([-1, 0])
    if i % 2:
        v[1] = -v[1]
    return lambda t: screenpos + 800 * d(t - 0.05 * i) * rotMatrix(-0.2 * d(t) * a).dot(v)

def slide_in_left(screenpos, i, nletters):
    scale = 2
    v = np.array([-300, 0])
    d = lambda t: min(1, max(0, 3 * (t * scale)))
    return lambda t: screenpos + v * (1 - d(t))

def moveLetters(letters, funcpos):
    return [letter.set_pos(funcpos(letter.screenpos, i, len(letters))) for i, letter in enumerate(letters)]

effects = [vortex, cascade, arrive, vortexout, slide_in_left]
audio_clip = AudioFileClip("source_audio.mp3")
clip_duration = audio_clip.duration / len(image_files)

clips = []
total_image_duration = 0

for index, image_file in enumerate(image_files):
    img_clip = ImageClip(image_file).set_duration(fixed_image_duration).fx(resize, newsize=screensize)
    word = words[index]
    txtClip = TextClip(word, color="white", font="Amiri-Bold", kerning=5, fontsize=200, method='caption', size=(screensize[0]*0.9, None)).set_duration(fixed_image_duration)
    cvc = CompositeVideoClip([txtClip.set_pos("center")], size=screensize)
    letters = findObjects(cvc)
    animated_clips = moveLetters(letters, random.choice(effects))
    animated_clip = CompositeVideoClip(animated_clips, size=screensize).subclip(0, fixed_image_duration)
    video = CompositeVideoClip([img_clip, animated_clip.set_position("center")])
    clips.append(video)
    total_image_duration += fixed_image_duration

final_clip = concatenate_videoclips(clips, method="compose", padding=-1)
final_clip = final_clip.set_audio(audio_clip.set_duration(total_image_duration))
final_clip.write_videofile("ready_video.mp4", fps=24, codec="libx264", audio_codec="aac")

print("Video created successfully!")
