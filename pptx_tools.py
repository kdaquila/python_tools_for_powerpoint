import pptx


def load_pptx_file(pptx_file_path):
    # Load existing presentation
    try:
        prs = pptx.Presentation(pptx_file_path)
        print("Found existing file, will copy and edit")
        return prs
    except pptx.exc.PackageNotFoundError:
        print("No existing file found")
        return None


