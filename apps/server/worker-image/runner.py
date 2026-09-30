import ctypes
import ctypes.util
import errno
import os
import stat
from pathlib import Path

# Install the filter before evaluating untrusted source. Filters inherit across forks.
libseccomp = ctypes.CDLL(ctypes.util.find_library("seccomp"), use_errno=True)
libseccomp.seccomp_init.argtypes = [ctypes.c_uint32]
libseccomp.seccomp_init.restype = ctypes.c_void_p
libseccomp.seccomp_syscall_resolve_name.argtypes = [ctypes.c_char_p]
libseccomp.seccomp_syscall_resolve_name.restype = ctypes.c_int
libseccomp.seccomp_rule_add.argtypes = [ctypes.c_void_p, ctypes.c_uint32, ctypes.c_int, ctypes.c_uint]
libseccomp.seccomp_load.argtypes = [ctypes.c_void_p]
libseccomp.seccomp_release.argtypes = [ctypes.c_void_p]
filter_context = libseccomp.seccomp_init(0x7FFF0000)  # SCMP_ACT_ALLOW
if not filter_context:
    raise RuntimeError("seccomp unavailable")
try:
    for syscall in (b"execve", b"execveat"):
        number = libseccomp.seccomp_syscall_resolve_name(syscall)
        if number < 0 or libseccomp.seccomp_rule_add(filter_context, 0x00050000 | errno.EPERM, number, 0):
            raise RuntimeError("seccomp rule unavailable")
    if libseccomp.seccomp_load(filter_context):
        raise RuntimeError("seccomp installation failed")
finally:
    libseccomp.seccomp_release(filter_context)

output_fd = os.dup(1)
os.dup2(os.open(os.devnull, os.O_WRONLY), 1)
source = Path("/input/source.py").read_bytes()
exec(compile(source, "/input/source.py", "exec"), {"__name__": "__main__", "__file__": "/input/source.py"})
docx_fd = os.open("/work/output.docx", os.O_RDONLY | os.O_NOFOLLOW)
try:
    result = os.fstat(docx_fd)
    if not stat.S_ISREG(result.st_mode) or not 0 < result.st_size <= 25 * 1024 * 1024:
        raise ValueError("Invalid document output")
    while data := os.read(docx_fd, 64 * 1024):
        view = memoryview(data)
        while view:
            view = view[os.write(output_fd, view):]
finally:
    os.close(docx_fd)
