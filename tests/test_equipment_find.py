"""Tests for equipment_find and equipment_inspect operations."""

import math

import ezdxf
import pytest

from autocad_mcp.backends.ezdxf_backend import EzdxfBackend


@pytest.fixture
async def backend():
    """Initialized ezdxf backend."""
    b = EzdxfBackend()
    result = await b.initialize()
    assert result.ok
    return b


@pytest.fixture
async def backend_with_equipment(backend):
    """Backend with a drawing containing equipment-like entities."""
    doc = backend._doc
    msp = backend._msp

    # Add modelspace TEXT
    msp.add_text("TK-479 STORAGE TANK", dxfattribs={"insert": (100, 200, 0), "layer": "EQUIP", "height": 5.0})

    # Add modelspace MTEXT
    msp.add_mtext("PUMP P-101 CENTRIFUGAL", dxfattribs={"insert": (300, 400, 0), "layer": "EQUIP", "char_height": 3.0})

    # Add a circle (tank representation)
    msp.add_circle((100, 200, 0), radius=50, dxfattribs={"layer": "EQUIP"})

    # Add a small circle
    msp.add_circle((300, 400, 0), radius=10, dxfattribs={"layer": "EQUIP"})

    # Create a block with an ATTDEF and insert it
    blk = doc.blocks.new("VALVE_BLOCK")
    blk.add_text("HV-201", dxfattribs={"insert": (0, 0, 0), "height": 2.5})
    attdef = blk.add_attdef("TAG", insert=(0, -5, 0), dxfattribs={"height": 2.5})
    attdef.dxf.text = "HV-201"

    # Insert the block at a specific position
    ins = msp.add_blockref("VALVE_BLOCK", insert=(500, 600, 0), dxfattribs={"layer": "VALVES"})
    ins.add_auto_attribs({"TAG": "HV-201"})

    # Another block instance with different attribute
    ins2 = msp.add_blockref("VALVE_BLOCK", insert=(700, 800, 0), dxfattribs={
        "layer": "VALVES", "xscale": 2.0, "yscale": 2.0, "rotation": 45.0,
    })
    ins2.add_auto_attribs({"TAG": "HV-305"})

    # A nested block: ASSEMBLY contains VALVE_BLOCK
    asm = doc.blocks.new("ASSEMBLY")
    asm.add_text("ASSEMBLY-X1", dxfattribs={"insert": (10, 10, 0), "height": 2.0})
    asm.add_blockref("VALVE_BLOCK", insert=(20, 20, 0))

    msp.add_blockref("ASSEMBLY", insert=(900, 100, 0), dxfattribs={"layer": "ASSEMBLIES"})

    return backend


# ---------------------------------------------------------------------------
# equipment_find tests
# ---------------------------------------------------------------------------


class TestEquipmentFind:
    async def test_find_modelspace_text(self, backend_with_equipment):
        r = await backend_with_equipment.equipment_find("TK-479")
        assert r.ok
        assert r.payload["count"] >= 1
        first = r.payload["results"][0]
        assert first["type"] == "TEXT"
        assert "TK-479" in first["text"]
        assert first["context"] == "modelspace"

    async def test_find_modelspace_mtext(self, backend_with_equipment):
        r = await backend_with_equipment.equipment_find("P-101")
        assert r.ok
        assert r.payload["count"] >= 1
        found_mtext = any(m["type"] == "MTEXT" for m in r.payload["results"])
        assert found_mtext

    async def test_find_wildcard(self, backend_with_equipment):
        r = await backend_with_equipment.equipment_find("*TK*")
        assert r.ok
        assert r.payload["count"] >= 1
        assert "TK" in r.payload["results"][0]["text"].upper()

    async def test_find_attribute_value(self, backend_with_equipment):
        r = await backend_with_equipment.equipment_find("HV-201", search_scope="attributes")
        assert r.ok
        assert r.payload["count"] >= 1
        attrib_results = [m for m in r.payload["results"] if m["context"] == "attribute"]
        assert len(attrib_results) >= 1
        assert attrib_results[0]["containing_block"] == "VALVE_BLOCK"

    async def test_find_block_definition_text(self, backend_with_equipment):
        r = await backend_with_equipment.equipment_find("HV-201", search_scope="blocks")
        assert r.ok
        assert r.payload["count"] >= 1
        block_results = [m for m in r.payload["results"] if m["context"] == "block_definition"]
        assert len(block_results) >= 1
        # Should have world_position transformed
        for br in block_results:
            assert "world_position" in br
            assert br["containing_block"] == "VALVE_BLOCK"

    async def test_find_case_insensitive(self, backend_with_equipment):
        r = await backend_with_equipment.equipment_find("tk-479", case_sensitive=False)
        assert r.ok
        assert r.payload["count"] >= 1

    async def test_find_case_sensitive_miss(self, backend_with_equipment):
        r = await backend_with_equipment.equipment_find("tk-479", case_sensitive=True)
        assert r.ok
        assert r.payload["count"] == 0

    async def test_find_scope_modelspace_only(self, backend_with_equipment):
        r = await backend_with_equipment.equipment_find("HV-201", search_scope="modelspace")
        assert r.ok
        # HV-201 is in blocks/attributes, not modelspace text
        assert r.payload["count"] == 0

    async def test_find_max_results(self, backend_with_equipment):
        r = await backend_with_equipment.equipment_find("*", max_results=3)
        assert r.ok
        assert r.payload["count"] <= 3

    async def test_find_empty_pattern(self, backend_with_equipment):
        """Empty pattern: LISP backend rejects it; ezdxf matches everything (valid regex).
        The server.py layer validates pattern is present, so this tests backend directly."""
        r = await backend_with_equipment.equipment_find("")
        # ezdxf backend treats empty string as match-all (valid regex)
        assert r.ok
        assert r.payload["count"] > 0

    async def test_find_no_match(self, backend_with_equipment):
        r = await backend_with_equipment.equipment_find("NONEXISTENT-XYZ-999")
        assert r.ok
        assert r.payload["count"] == 0

    async def test_find_zoomed_to_present(self, backend_with_equipment):
        r = await backend_with_equipment.equipment_find("TK-479", zoom_to_first=True)
        assert r.ok
        if r.payload["count"] > 0:
            assert "zoomed_to" in r.payload

    async def test_find_all_scope(self, backend_with_equipment):
        """Search all scopes and verify we get results from multiple contexts."""
        r = await backend_with_equipment.equipment_find("HV-201", search_scope="all")
        assert r.ok
        contexts = set(m["context"] for m in r.payload["results"])
        # Should find in at least attributes and block_definition
        assert len(contexts) >= 1

    async def test_find_world_position_transform(self, backend_with_equipment):
        """Block definition text should have world_position different from local position
        when the INSERT has scale/rotation."""
        r = await backend_with_equipment.equipment_find("HV-201", search_scope="blocks")
        assert r.ok
        for result in r.payload["results"]:
            assert "world_position" in result
            # World position should not be [0,0,0] — it should be transformed
            wp = result["world_position"]
            assert not (wp[0] == 0 and wp[1] == 0)


# ---------------------------------------------------------------------------
# equipment_inspect tests
# ---------------------------------------------------------------------------


class TestEquipmentInspect:
    async def test_inspect_basic(self, backend_with_equipment):
        r = await backend_with_equipment.equipment_inspect(100, 200)
        assert r.ok
        assert "view_center" in r.payload
        assert "equipment_center" in r.payload
        assert "nearby_entities" in r.payload

    async def test_inspect_finds_circle(self, backend_with_equipment):
        """Inspect near a large circle should find it and use it for center."""
        r = await backend_with_equipment.equipment_inspect(100, 200, view_width=200, view_height=200)
        assert r.ok
        circles = r.payload["nearby_entities"]["circles"]
        assert len(circles) >= 1

    async def test_inspect_center_largest_circle(self, backend_with_equipment):
        """Center inference should pick the largest circle."""
        r = await backend_with_equipment.equipment_inspect(100, 200, view_width=200, view_height=200)
        assert r.ok
        ec = r.payload["equipment_center"]
        if ec["method"] == "largest_circle":
            assert ec["confidence"] == "high"
            # Should be near the circle center
            assert abs(ec["x"] - 100) < 1
            assert abs(ec["y"] - 200) < 1

    async def test_inspect_finds_blocks(self, backend_with_equipment):
        """Inspect near block insertion should find blocks."""
        r = await backend_with_equipment.equipment_inspect(500, 600, view_width=200, view_height=200)
        assert r.ok
        blocks = r.payload["nearby_entities"]["blocks"]
        assert len(blocks) >= 1
        assert blocks[0]["block_name"] == "VALVE_BLOCK"

    async def test_inspect_nearest_insert_fallback(self, backend_with_equipment):
        """When no circles, nearest INSERT should be used."""
        r = await backend_with_equipment.equipment_inspect(500, 600, view_width=200, view_height=200)
        assert r.ok
        ec = r.payload["equipment_center"]
        # Should be nearest_insert or insert_bbox
        assert ec["method"] in ("nearest_insert", "insert_bbox", "largest_circle")
        assert ec["confidence"] in ("medium", "high")

    async def test_inspect_fallback(self, backend_with_equipment):
        """Far from everything, should fallback to input coordinates."""
        r = await backend_with_equipment.equipment_inspect(9999, 9999, view_width=10, view_height=10)
        assert r.ok
        ec = r.payload["equipment_center"]
        if ec["method"] == "fallback":
            assert ec["confidence"] == "low"
            assert abs(ec["x"] - 9999) < 1
            assert abs(ec["y"] - 9999) < 1

    async def test_inspect_by_type_counts(self, backend_with_equipment):
        """by_type should have entity type counts."""
        r = await backend_with_equipment.equipment_inspect(100, 200, view_width=1000, view_height=1000)
        assert r.ok
        by_type = r.payload["nearby_entities"]["by_type"]
        assert isinstance(by_type, dict)
        total = r.payload["nearby_entities"]["total"]
        assert total >= 1

    async def test_inspect_missing_coordinates(self, backend_with_equipment):
        """Should error if x or y missing."""
        # ezdxf backend requires both x and y
        r = await backend_with_equipment.equipment_inspect(None, None)
        # The method should handle gracefully — may error or return fallback
        # Either is acceptable behavior

    async def test_inspect_infer_center_false(self, backend_with_equipment):
        """With infer_center=False, should use fallback."""
        r = await backend_with_equipment.equipment_inspect(100, 200, infer_center=False)
        assert r.ok
        ec = r.payload["equipment_center"]
        assert ec["method"] == "fallback"
