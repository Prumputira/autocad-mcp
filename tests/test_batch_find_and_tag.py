"""Tests for batch_find_and_tag on the ezdxf backend."""

import pytest

from autocad_mcp.backends.ezdxf_backend import EzdxfBackend


@pytest.fixture
async def backend():
    """Initialized ezdxf backend with sample entities."""
    b = EzdxfBackend()
    result = await b.initialize()
    assert result.ok
    return b


@pytest.fixture
async def populated_backend(backend):
    """Backend with several known text entities for searching."""
    # Add TEXT entities that the batch can find
    await backend.create_text(100, 200, "FT-575", height=8.0, layer="E-ANNO-TEXT")
    await backend.create_text(300, 400, "BS-564", height=8.0, layer="E-ANNO-TEXT")
    await backend.create_text(500, 600, "FT-577", height=8.0, layer="E-ANNO-TEXT")
    # Add an MTEXT entity
    await backend.create_mtext(700, 800, 100, "PMP-101", height=8.0, layer="E-ANNO-TEXT")
    return backend


class TestBatchFindAndTag:
    async def test_all_found(self, populated_backend):
        """All tags exist in drawing — all should be placed."""
        r = await populated_backend.batch_find_and_tag(
            ["FT-575", "BS-564", "FT-577"],
        )
        assert r.ok
        assert r.payload["placed"] == 3
        assert r.payload["not_found"] == []
        assert len(r.payload["results"]) == 3
        for entry in r.payload["results"]:
            assert entry["status"] == "placed"
            assert "cube_handle" in entry
            assert "leader_handle" in entry
            assert "mtext_handle" in entry
            assert "line_handle" in entry

    async def test_mixed_found_and_not_found(self, populated_backend):
        """Some tags found, some not."""
        r = await populated_backend.batch_find_and_tag(
            ["FT-575", "NONEXISTENT-1", "BS-564", "NONEXISTENT-2"],
        )
        assert r.ok
        assert r.payload["placed"] == 2
        assert set(r.payload["not_found"]) == {"NONEXISTENT-1", "NONEXISTENT-2"}
        assert len(r.payload["results"]) == 4

    async def test_empty_tag_list(self, backend):
        """Empty list should succeed with zero placed."""
        r = await backend.batch_find_and_tag([])
        assert r.ok
        assert r.payload["placed"] == 0
        assert r.payload["not_found"] == []
        assert r.payload["results"] == []

    async def test_single_tag(self, populated_backend):
        """Single tag should work fine."""
        r = await populated_backend.batch_find_and_tag(["PMP-101"])
        assert r.ok
        assert r.payload["placed"] == 1
        assert r.payload["not_found"] == []
        assert len(r.payload["results"]) == 1
        assert r.payload["results"][0]["status"] == "placed"

    async def test_all_not_found(self, backend):
        """No tags exist — all should be not_found."""
        r = await backend.batch_find_and_tag(
            ["NOPE-1", "NOPE-2", "NOPE-3"],
        )
        assert r.ok
        assert r.payload["placed"] == 0
        assert len(r.payload["not_found"]) == 3
        for entry in r.payload["results"]:
            assert entry["status"] == "not_found"

    async def test_custom_parameters(self, populated_backend):
        """Cube size, direction, text height should be passed through."""
        r = await populated_backend.batch_find_and_tag(
            ["FT-575"],
            cube_size=12.0,
            direction="left",
            text_height=6.0,
        )
        assert r.ok
        assert r.payload["placed"] == 1

    async def test_no_document(self):
        """Should fail gracefully with no document open."""
        b = EzdxfBackend()
        r = await b.batch_find_and_tag(["FT-575"])
        assert not r.ok
        assert "No document" in r.error

    async def test_result_has_position(self, populated_backend):
        """Each placed result should include the position where it was found."""
        r = await populated_backend.batch_find_and_tag(["FT-575"])
        assert r.ok
        entry = r.payload["results"][0]
        assert "position" in entry
        pos = entry["position"]
        # FT-575 was placed at (100, 200)
        assert abs(pos[0] - 100) < 1.0
        assert abs(pos[1] - 200) < 1.0
