@echo off
for %%A in (*.wmv) do (
  ffmpeg -i "%%A" -vcodec libx264 -crf 25 -preset medium -tune stillimage -profile:v baseline -level 3.0 -vf "fps=30000/1001,format=yuv420p" -acodec aac -ab 96k -ac 2 -ar 44100 -absf aac_adtstoasc -async 1 "%%~nA".mp4
)
exit /b 0