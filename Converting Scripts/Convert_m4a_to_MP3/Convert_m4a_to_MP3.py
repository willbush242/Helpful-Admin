import subprocess
from pathlib import Path
import imageio_ffmpeg

root_folder = Path(r"path/to/audio")

ffmpeg = imageio_ffmpeg.get_ffmpeg_exe()

m4a_files = list(root_folder.rglob("*.m4a"))

print(f"Found {len(m4a_files)} .m4a file(s).\n")

for input_file in m4a_files:
    output_file = input_file.with_suffix(".mp3")

    print(f"Converting:")
    print(f"  {input_file}")
    print(f"  -> {output_file}")

    result = subprocess.run(
        [
            ffmpeg,
            "-y",
            "-i", str(input_file),
            "-vn",
            "-c:a", "libmp3lame",
            "-q:a", "0",
            str(output_file)
        ],
        capture_output=True,
        text=True
    )

    if (
        result.returncode == 0
        and output_file.exists()
        and output_file.stat().st_size > 0
    ):
        print("  Conversion successful.")

        try:
            input_file.unlink()
            print("  Original .m4a deleted.\n")
        except OSError as e:
            print("  MP3 created, but couldn't delete original:")
            print(f"  {e}\n")

    else:
        print("  Conversion FAILED.")
        print("  Original .m4a has NOT been deleted.")
        print(result.stderr)
        print()

print("All files processed.")
