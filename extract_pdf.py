"""
Minimal pure-Python PDF text extractor (no external deps).
Extracts text from PDF content streams by:
1. Finding stream...endstream blocks
2. Decompressing FlateDecode streams with zlib
3. Parsing text-showing operators (Tj, TJ, ', ")
"""
import re
import sys
import zlib


def extract_pdf_text(path: str) -> str:
    with open(path, "rb") as f:
        data = f.read()

    # Find all objects
    out_lines = []

    # Find all stream blocks
    # Pattern: <<...>>\nstream\n...\nendstream
    stream_re = re.compile(rb"<<(?P<dict>.*?)>>\s*stream\s*\r?\n(?P<body>.*?)\r?\nendstream", re.DOTALL)

    for m in stream_re.finditer(data):
        d = m.group("dict")
        body = m.group("body")

        # Check if FlateDecode
        if b"/FlateDecode" in d or b"/Fl " in d or b"/Fl\n" in d or b"/Fl/" in d:
            try:
                body = zlib.decompress(body)
            except Exception:
                # Try removing trailing garbage
                try:
                    body = zlib.decompressobj().decompress(body)
                except Exception:
                    continue
        elif b"/Filter" in d:
            # Other filters not supported; skip
            continue

        # Now parse text operators in the body
        text = parse_content_stream(body)
        if text:
            out_lines.append(text)

    return "\n".join(out_lines)


def parse_content_stream(body: bytes) -> str:
    """Parse PDF content stream and extract text strings."""
    text_chunks = []
    i = 0
    n = len(body)

    # We'll scan token by token. Look for:
    #   (literal string) Tj
    #   (literal string) '
    #   (a) (b) "       (uses two args, last is text)
    #   [(s1) -120 (s2)] TJ
    #   <hex string> Tj   (hex)

    while i < n:
        c = body[i:i+1]
        if c == b"(":
            # parse literal string
            s, j = parse_literal_string(body, i)
            i = j
            # skip whitespace
            while i < n and body[i:i+1] in b" \t\r\n":
                i += 1
            # check operator
            op = read_op(body, i)
            if op in (b"Tj", b"'"):
                text_chunks.append(decode_pdf_string(s))
                text_chunks.append("\n")
            elif op == b'"':
                text_chunks.append(decode_pdf_string(s))
                text_chunks.append("\n")
            else:
                text_chunks.append(decode_pdf_string(s))
            i += len(op)
        elif c == b"<":
            # could be hex string or dict
            if body[i:i+2] == b"<<":
                # dict, skip to >>
                depth = 1
                i += 2
                while i < n and depth > 0:
                    if body[i:i+2] == b"<<":
                        depth += 1
                        i += 2
                    elif body[i:i+2] == b">>":
                        depth -= 1
                        i += 2
                    else:
                        i += 1
            else:
                # hex string
                end = body.find(b">", i)
                if end == -1:
                    break
                hex_content = body[i+1:end].replace(b" ", b"").replace(b"\n", b"").replace(b"\r", b"")
                try:
                    s = bytes.fromhex(hex_content.decode("ascii"))
                    text_chunks.append(decode_pdf_string(s))
                except Exception:
                    pass
                i = end + 1
        elif c == b"[":
            # array — used by TJ. Extract any strings inside.
            depth = 1
            i += 1
            arr_text = []
            while i < n and depth > 0:
                cc = body[i:i+1]
                if cc == b"(":
                    s, j = parse_literal_string(body, i)
                    arr_text.append(decode_pdf_string(s))
                    i = j
                elif cc == b"<":
                    end = body.find(b">", i)
                    if end == -1:
                        break
                    hex_content = body[i+1:end].replace(b" ", b"").replace(b"\n", b"").replace(b"\r", b"")
                    try:
                        s = bytes.fromhex(hex_content.decode("ascii"))
                        arr_text.append(decode_pdf_string(s))
                    except Exception:
                        pass
                    i = end + 1
                elif cc == b"[":
                    depth += 1
                    i += 1
                elif cc == b"]":
                    depth -= 1
                    i += 1
                else:
                    i += 1
            text_chunks.append("".join(arr_text))
            text_chunks.append("\n")
        else:
            i += 1

    return "".join(text_chunks)


def parse_literal_string(data: bytes, start: int) -> tuple[bytes, int]:
    """Parse a literal string starting at data[start]==b'('. Return (content, next_index)."""
    assert data[start:start+1] == b"("
    i = start + 1
    depth = 1
    out = bytearray()
    while i < len(data) and depth > 0:
        c = data[i:i+1]
        if c == b"\\":
            # escape
            nxt = data[i+1:i+2]
            if nxt == b"n":
                out.append(0x0a); i += 2
            elif nxt == b"r":
                out.append(0x0d); i += 2
            elif nxt == b"t":
                out.append(0x09); i += 2
            elif nxt == b"b":
                out.append(0x08); i += 2
            elif nxt == b"f":
                out.append(0x0c); i += 2
            elif nxt in (b"(", b")", b"\\"):
                out.extend(nxt); i += 2
            elif nxt in b"01234567":
                # octal up to 3 digits
                octdigs = nxt
                i += 2
                for _ in range(2):
                    if i < len(data) and data[i:i+1] in b"01234567":
                        octdigs += data[i:i+1]
                        i += 1
                    else:
                        break
                out.append(int(octdigs, 8) & 0xff)
            elif nxt in (b"\n", b"\r"):
                # line continuation
                i += 2
                if data[i:i+1] == b"\n":
                    i += 1
            else:
                i += 2  # unknown escape
        elif c == b"(":
            depth += 1
            out.extend(c); i += 1
        elif c == b")":
            depth -= 1
            if depth == 0:
                i += 1
                break
            out.extend(c); i += 1
        else:
            out.extend(c); i += 1
    return bytes(out), i


def read_op(data: bytes, i: int) -> bytes:
    """Read an operator name at position i."""
    j = i
    while j < len(data) and data[j:j+1] not in b" \t\r\n([<":
        j += 1
        if j - i > 5:  # operators are short
            break
    return data[i:j]


def decode_pdf_string(b: bytes) -> str:
    """Decode PDF string bytes. Try UTF-16BE BOM, then Latin-1."""
    if b.startswith(b"\xfe\xff"):
        try:
            return b[2:].decode("utf-16-be", errors="replace")
        except Exception:
            pass
    # Default: treat as Latin-1 / PDFDocEncoding (close enough for ASCII)
    try:
        return b.decode("latin-1", errors="replace")
    except Exception:
        return ""


if __name__ == "__main__":
    path = sys.argv[1] if len(sys.argv) > 1 else "file ni pogi na naka a4.pdf"
    text = extract_pdf_text(path)
    print(text)
