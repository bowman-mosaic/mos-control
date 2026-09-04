"""No-hardware tests for recorded TIFF viewer caching."""

import numpy as np
import pytest

tifffile = pytest.importorskip("tifffile")

from modules import coolsnap


@pytest.fixture(autouse=True)
def _empty_tiff_cache():
    coolsnap._clear_tiff_cache()
    yield
    coolsnap._clear_tiff_cache()


def test_repeated_frame_requests_reuse_cached_source(tmp_path, monkeypatch):
    path = tmp_path / "video.tif"
    stack = np.arange(3 * 6 * 7, dtype=np.uint16).reshape(3, 6, 7)
    tifffile.imwrite(path, stack, photometric="minisblack")

    real_memmap = coolsnap.tifffile.memmap
    calls = []

    def counting_memmap(*args, **kwargs):
        calls.append(args[0])
        return real_memmap(*args, **kwargs)

    monkeypatch.setattr(coolsnap.tifffile, "memmap", counting_memmap)

    frame_1, idx_1, count_1, shape_1 = coolsnap._get_tiff_frame(path, 1)
    frame_2, idx_2, count_2, shape_2 = coolsnap._get_tiff_frame(path, 2)

    assert len(calls) == 1
    np.testing.assert_array_equal(frame_1, stack[1])
    np.testing.assert_array_equal(frame_2, stack[2])
    assert (idx_1, idx_2) == (1, 2)
    assert (count_1, count_2) == (3, 3)
    assert shape_1 == shape_2 == stack.shape


@pytest.mark.parametrize(
    ("data", "requested_idx", "expected_idx", "expected_count", "expected"),
    [
        (
            np.arange(20, dtype=np.uint16).reshape(4, 5),
            3,
            0,
            1,
            np.arange(20, dtype=np.uint16).reshape(4, 5),
        ),
        (
            np.arange(40, dtype=np.uint16).reshape(2, 4, 5),
            1,
            1,
            2,
            np.arange(40, dtype=np.uint16).reshape(2, 4, 5)[1],
        ),
    ],
    ids=["single", "grayscale-stack"],
)
def test_frame_selection_preserves_tiff_shape_semantics(
    tmp_path, data, requested_idx, expected_idx, expected_count, expected
):
    path = tmp_path / f"{data.ndim}d-{data.shape[-1]}.tif"
    tifffile.imwrite(path, data)

    frame, selected_idx, n_frames, shape = coolsnap._get_tiff_frame(
        path, requested_idx
    )

    np.testing.assert_array_equal(frame, expected)
    assert selected_idx == expected_idx
    assert n_frames == expected_count
    assert shape == data.shape


def test_three_dimensional_rgb_tiff_keeps_frame_stack_semantics(tmp_path):
    path = tmp_path / "rgb-metadata.tif"
    data = np.arange(4 * 5 * 3, dtype=np.uint8).reshape(4, 5, 3)
    tifffile.imwrite(path, data, photometric="rgb")

    frame, selected_idx, n_frames, shape = coolsnap._get_tiff_frame(path, 2)

    np.testing.assert_array_equal(frame, data[2])
    assert selected_idx == 2
    assert n_frames == 4
    assert shape == data.shape


def test_four_dimensional_tiff_remains_unsupported(tmp_path):
    path = tmp_path / "rgb-stack.tif"
    data = np.zeros((2, 4, 5, 3), dtype=np.uint8)
    tifffile.imwrite(path, data, photometric="rgb")

    with pytest.raises(coolsnap._UnsupportedTiffShape):
        coolsnap._get_tiff_frame(path, 0)


def test_preview_uses_local_pseudo_color_without_mutating_globals(
    tmp_path, monkeypatch
):
    path = tmp_path / "preview.tif"
    frame = np.array([[0, 21845], [43690, 65535]], dtype=np.uint16)
    tifffile.imwrite(path, frame, photometric="minisblack")
    encoded = []

    class FakeCv2:
        COLOR_RGB2BGR = 1
        IMWRITE_JPEG_QUALITY = 2

        @staticmethod
        def cvtColor(image, _code):
            return image[..., ::-1]

        @staticmethod
        def imencode(_extension, image, _options):
            encoded.append(image.copy())
            return True, np.array([1, 2, 3], dtype=np.uint8)

    monkeypatch.setattr(coolsnap, "_save_dir", str(tmp_path))
    monkeypatch.setattr(coolsnap, "_get_capture_meta",
                        lambda _path: ("green", None))
    monkeypatch.setattr(coolsnap, "_HAS_CV2", True)
    monkeypatch.setattr(coolsnap, "cv2", FakeCv2)
    monkeypatch.setattr(coolsnap, "_pseudo_color", (255, 0, 0))
    monkeypatch.setattr(coolsnap, "_pseudo_color_name", "red")

    result = coolsnap.cam_npy_preview(
        path.name, display_min=0, display_max=65535
    )

    assert result["color"] == "green"
    assert len(encoded) == 1
    np.testing.assert_array_equal(encoded[0][..., 0], 0)
    np.testing.assert_array_equal(
        encoded[0][..., 1], np.array([[0, 85], [170, 255]], dtype=np.uint8)
    )
    np.testing.assert_array_equal(encoded[0][..., 2], 0)
    assert coolsnap._pseudo_color == (255, 0, 0)
    assert coolsnap._pseudo_color_name == "red"


def test_replacing_file_invalidates_cached_source(tmp_path, monkeypatch):
    path = tmp_path / "replace.tif"
    original = np.full((2, 5, 6), 11, dtype=np.uint16)
    replacement = np.full((3, 5, 6), 29, dtype=np.uint16)
    tifffile.imwrite(path, original, photometric="minisblack")

    real_memmap = coolsnap.tifffile.memmap
    calls = []

    def counting_memmap(*args, **kwargs):
        calls.append(args[0])
        return real_memmap(*args, **kwargs)

    monkeypatch.setattr(coolsnap.tifffile, "memmap", counting_memmap)

    first, _, _, _ = coolsnap._get_tiff_frame(path, 0)
    # Windows does not permit replacing an actively mapped file. Close the
    # mapping without removing its cache entry to preserve the stale-entry
    # condition that the signature check must detect.
    cached = next(iter(coolsnap._tiff_cache.values()))
    coolsnap._close_tiff_array(cached.array)
    tifffile.imwrite(path, replacement, photometric="minisblack")
    second, _, _, _ = coolsnap._get_tiff_frame(path, 0)

    np.testing.assert_array_equal(first, original[0])
    np.testing.assert_array_equal(second, replacement[0])
    assert len(calls) == 2


def test_capture_listing_uses_tiff_metadata_only(tmp_path, monkeypatch):
    path = tmp_path / "listed.tif"
    stack = np.zeros((3, 4, 5), dtype=np.uint16)
    tifffile.imwrite(path, stack, photometric="minisblack")
    monkeypatch.setattr(coolsnap, "_base_save_dir", str(tmp_path))

    def unexpected_decode(*args, **kwargs):
        raise AssertionError("capture listing decoded TIFF pixels")

    monkeypatch.setattr(coolsnap.tifffile, "memmap", unexpected_decode)
    monkeypatch.setattr(coolsnap.tifffile, "imread", unexpected_decode)

    assert coolsnap.cam_list_captures() == {
        "files": [{"name": path.name, "shape": [3, 4, 5]}]
    }


def test_explicit_pseudo_color_does_not_mutate_global_state():
    frame = np.arange(20, dtype=np.uint16).reshape(4, 5)
    previous = coolsnap._pseudo_color
    coolsnap._pseudo_color = (255, 0, 0)
    try:
        grayscale, _, _ = coolsnap._apply_bcg(frame, pseudo_color=None)
        tinted, _, _ = coolsnap._apply_bcg(
            frame, pseudo_color=(0, 255, 0)
        )
        assert grayscale.ndim == 2
        assert tinted.shape == (4, 5, 3)
        assert coolsnap._pseudo_color == (255, 0, 0)
    finally:
        coolsnap._pseudo_color = previous
