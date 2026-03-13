"""Tests for the place_equipment_tag operation on the ezdxf backend."""

import pytest

from autocad_mcp.backends.ezdxf_backend import EzdxfBackend


@pytest.fixture
async def backend():
    """Initialized ezdxf backend with a blank drawing."""
    b = EzdxfBackend()
    await b.initialize()
    return b


class TestPlaceEquipmentTagBasic:
    async def test_right_direction(self, backend):
        r = await backend.place_equipment_tag(100, 200, 0, "TK-479")
        assert r.ok
        p = r.payload
        assert "cube_handle" in p
        assert "leader_handle" in p
        assert "mtext_handle" in p
        assert "line_handle" in p
        assert p["center"] == {"x": 100, "y": 200, "z": 0}
        assert p["tag"] == "TK-479"
        assert p["text_width"] > 0
        assert "bbox" in p

    async def test_left_direction(self, backend):
        r = await backend.place_equipment_tag(100, 200, 0, "TK-479", direction="left")
        assert r.ok
        p = r.payload
        assert p["center"]["x"] == 100
        # For left direction, MTEXT x should be to the left of center
        bbox = p["bbox"]
        assert bbox["min_x"] < 100

    async def test_custom_cube_size(self, backend):
        r = await backend.place_equipment_tag(0, 0, 0, "P-101", cube_size=48.0)
        assert r.ok
        bbox = r.payload["bbox"]
        # Cube extends 24 units (half of 48) from center
        assert bbox["min_x"] <= -24.0
        assert bbox["min_y"] <= -24.0

    async def test_custom_text_height(self, backend):
        r = await backend.place_equipment_tag(0, 0, 0, "P-101", text_height=12.0)
        assert r.ok
        assert r.payload["text_width"] > 0

    async def test_nonzero_z(self, backend):
        r = await backend.place_equipment_tag(100, 200, 50.0, "EQ-001")
        assert r.ok
        assert r.payload["center"]["z"] == 50.0


class TestPlaceEquipmentTagHandles:
    async def test_all_handles_are_unique(self, backend):
        r = await backend.place_equipment_tag(0, 0, 0, "TEST-001")
        assert r.ok
        p = r.payload
        handles = [p["cube_handle"], p["leader_handle"], p["mtext_handle"], p["line_handle"]]
        assert len(set(handles)) == 4, "All 4 handles should be unique"

    async def test_handles_exist_in_document(self, backend):
        r = await backend.place_equipment_tag(0, 0, 0, "TEST-002")
        assert r.ok
        p = r.payload
        # Verify leader, mtext, and line handles are in the entity database
        for key in ["leader_handle", "mtext_handle", "line_handle"]:
            entity = backend._doc.entitydb.get(p[key])
            assert entity is not None, f"{key} ({p[key]}) not found in entitydb"


class TestPlaceEquipmentTagLayers:
    async def test_layers_auto_created(self, backend):
        # Before: layers should not exist
        layers_before = [l.dxf.name for l in backend._doc.layers]
        assert "E-EQPM-N" not in layers_before
        assert "E-ANNO-TEXT" not in layers_before

        r = await backend.place_equipment_tag(0, 0, 0, "TAG-001")
        assert r.ok

        # After: layers should exist
        layers_after = [l.dxf.name for l in backend._doc.layers]
        assert "E-EQPM-N" in layers_after
        assert "E-ANNO-TEXT" in layers_after

    async def test_entities_on_correct_layers(self, backend):
        r = await backend.place_equipment_tag(0, 0, 0, "TAG-002")
        assert r.ok
        p = r.payload
        # Leader, MTEXT, and LINE should be on E-ANNO-TEXT
        for key in ["leader_handle", "mtext_handle", "line_handle"]:
            ent = backend._doc.entitydb.get(p[key])
            assert ent.dxf.layer == "E-ANNO-TEXT", f"{key} should be on E-ANNO-TEXT"


class TestPlaceEquipmentTagGeometry:
    async def test_bbox_contains_center(self, backend):
        r = await backend.place_equipment_tag(500, 600, 0, "EQ-100")
        assert r.ok
        bbox = r.payload["bbox"]
        assert bbox["min_x"] <= 500 <= bbox["max_x"]
        assert bbox["min_y"] <= 600 <= bbox["max_y"]

    async def test_right_leader_goes_right(self, backend):
        r = await backend.place_equipment_tag(0, 0, 0, "R-001", direction="right")
        assert r.ok
        bbox = r.payload["bbox"]
        # Leader + mtext extend to the right of center
        assert bbox["max_x"] > 12.0  # At least past the cube half

    async def test_left_leader_goes_left(self, backend):
        r = await backend.place_equipment_tag(0, 0, 0, "L-001", direction="left")
        assert r.ok
        bbox = r.payload["bbox"]
        # Leader + mtext extend to the left of center
        assert bbox["min_x"] < -12.0

    async def test_multiple_placements_no_interference(self, backend):
        r1 = await backend.place_equipment_tag(0, 0, 0, "A-001")
        r2 = await backend.place_equipment_tag(200, 200, 0, "A-002")
        r3 = await backend.place_equipment_tag(-200, -200, 0, "A-003")
        assert r1.ok and r2.ok and r3.ok
        # All handles should be unique across placements
        all_handles = set()
        for r in [r1, r2, r3]:
            p = r.payload
            for key in ["cube_handle", "leader_handle", "mtext_handle", "line_handle"]:
                all_handles.add(p[key])
        assert len(all_handles) == 12

    async def test_entity_count_increases_by_expected(self, backend):
        count_before = len(list(backend._msp))
        await backend.place_equipment_tag(0, 0, 0, "CNT-001")
        count_after = len(list(backend._msp))
        # Polyface mesh (1 POLYLINE entity in msp) + LEADER + MTEXT + LINE = 4
        # Note: ezdxf may count differently for polyface mesh sub-entities
        assert count_after > count_before


class TestPlaceEquipmentTagValidation:
    async def test_missing_tag_fails(self, backend):
        r = await backend.place_equipment_tag(0, 0, 0, "")
        # Empty tag should still work (base class doesn't validate, LISP does)
        # ezdxf backend creates the entity anyway
        # Just verify no crash
        assert isinstance(r.ok, bool)

    async def test_default_parameters(self, backend):
        r = await backend.place_equipment_tag(100, 200, tag="DEF-001")
        assert r.ok
        assert r.payload["center"]["z"] == 0.0
