def extract_save_dir(file_path: str) -> str:
    if '/' in file_path:
        return file_path.rsplit('/', 1)[0]
    elif '\\' in file_path:  # Handle Windows paths
        return file_path.rsplit('\\', 1)[0]
    else:
        return '.'  # Current directory 