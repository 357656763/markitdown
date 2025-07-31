import sys

from typing import BinaryIO, Any

from ._html_converter import HtmlConverter
from ..converter_utils.docx.pre_process import pre_process_docx
from .._base_converter import DocumentConverterResult
from .._stream_info import StreamInfo
from .._exceptions import MissingDependencyException, MISSING_DEPENDENCY_MESSAGE

# Try loading optional (but in this case, required) dependencies
# Save reporting of any exceptions for later
_dependency_exc_info = None
try:
    import mammoth
except ImportError:
    # Preserve the error and stack trace for later
    _dependency_exc_info = sys.exc_info()


ACCEPTED_MIME_TYPE_PREFIXES = [
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
]

ACCEPTED_FILE_EXTENSIONS = [".docx"]


class DocxConverter(HtmlConverter):
    """
    Converts DOCX files to Markdown. Style information (e.g.m headings) and tables are preserved where possible.
    """

    def __init__(self):
        super().__init__()
        self._html_converter = HtmlConverter()

    def accepts(
        self,
        file_stream: BinaryIO,
        stream_info: StreamInfo,
        **kwargs: Any,  # Options to pass to the converter
    ) -> bool:
        mimetype = (stream_info.mimetype or "").lower()
        extension = (stream_info.extension or "").lower()

        if extension in ACCEPTED_FILE_EXTENSIONS:
            return True

        for prefix in ACCEPTED_MIME_TYPE_PREFIXES:
            if mimetype.startswith(prefix):
                return True

        return False

    def convert(
        self,
        file_stream: BinaryIO,
        stream_info: StreamInfo,
        **kwargs: Any,  # Options to pass to the converter
    ) -> DocumentConverterResult:
        # Check: the dependencies
        if _dependency_exc_info is not None:
            raise MissingDependencyException(
                MISSING_DEPENDENCY_MESSAGE.format(
                    converter=type(self).__name__,
                    extension=".docx",
                    feature="docx",
                )
            ) from _dependency_exc_info[
                1
            ].with_traceback(  # type: ignore[union-attr]
                _dependency_exc_info[2]
            )

        style_map = kwargs.get("style_map", None)
        # Developers customize image processing
        """
            eg:
            from mammoth.images import img_element
            @img_element
            def my_data_url(image):
                bucket_name = os.getenv("Bucket_Name", None)
                MINIO_URL = os.getenv("MINIO_URL", "127.0.0.1")
                
                with image.open() as image_bytes:
                    object_path = image_bytes.name
                    img_b = image_bytes.read()
                    suffix_with_dot = os.path.splitext(object_path)[1]  # 结果：.jpeg
                    object_name = hashlib.sha256(img_b).hexdigest() + suffix_with_dot
                    print(len(img_b), object_name)
                    minio_client.put_object(
                        bucket_name=bucket_name,
                        object_name=object_name,
                        data=BytesIO(img_b),
                        length=len(img_b)
                    )
                url = f"http://{MINIO_URL}/{bucket_name}/{object_name}"
            return {  "src": url  }
        """
        convert_image = kwargs.get("convert_image", None)

        pre_process_stream = pre_process_docx(file_stream)
        
        return self._html_converter.convert_string(
            mammoth.convert_to_html(pre_process_stream, style_map=style_map,convert_image=convert_image).value,
            **kwargs,
        )
