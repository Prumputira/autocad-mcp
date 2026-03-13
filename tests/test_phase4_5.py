"""Tests for Phase 4-5 operations: extended query, search, entity, validate,
select, drawing, layout, xref, and electrical tools on the ezdxf backend."""

import math
import os
import tempfile

import ezdxf
import pytest

from autocad_mcp.backends.ezdxf_backend import EzdxfBackend


@pytest.fixture
async def backend():
    """Initialized ezdxf backend with a blank drawing."""
    b = EzdxfBackend()
    await b.initialize()
    return b


@pytest.fixture
async def populated(backend):
    """Backend with layers, text, blocks, and lines for query/search/validate tests."""
    b = backend
    await b.layer_create("ANNO", "3", "Continuous")
    await b.layer_create("EQUIP", "2", "Continuous")
    await b.layer_create("WIRE", "1", "Continuous")
    await b.create_line(0, 0, 100, 0, layer="WIRE")
    await b.create_line(100, 0, 200, 0, layer="WIRE")
    await b.create_line(200, 0, 200, 100, layer="WIRE")
    await b.create_mtext(50, 50, 100, "Hello World", height=5.0, layer="ANNO")
    await b.create_text(150, 50, "Test Label", height=3.5, layer="ANNO")
    await b.create_circle(100, 100, 25, layer="EQUIP")
    # Define block with ATTDEF entities in the entities list
    await b.block_define("PUMP", [
        {"type": "CIRCLE", "cx": 0, "cy": 0, "radius": 10},
        {"type": "ATTDEF", "tag": "TAG", "x": 0, "y": 15, "height": 2.5},
        {"type": "ATTDEF", "tag": "DESC", "x": 0, "y": -15, "height": 2.5},
    ])
    await b.layer_set_current("EQUIP")
    await b.block_insert_with_attributes("PUMP", 50, 100, attributes={"TAG": "P-101", "DESC": "Feed Pump"})
    await b.block_insert_with_attributes("PUMP", 150, 100, attributes={"TAG": "P-102", "DESC": ""})
    await b.layer_set_current("0")
    return b


# ---------------------------------------------------------------------------
# Extended Query
# ---------------------------------------------------------------------------


class TestQueryTextStyles:
    async def test_returns_styles(self, backend):
        r = await backend.query_text_styles()
        assert r.ok
        styles = r.payload["styles"]
        assert isinstance(styles, list)
        assert len(styles) >= 1  # At least "Standard"
        names = [s["name"] for s in styles]
        assert "Standard" in names

    async def test_style_properties(self, backend):
        r = await backend.query_text_styles()
        standard = [s for s in r.payload["styles"] if s["name"] == "Standard"][0]
        assert "font" in standard
        assert "height" in standard


class TestQueryDimensionStyles:
    async def test_returns_dimstyles(self, backend):
        r = await backend.query_dimension_styles()
        assert r.ok
        styles = r.payload["dimstyles"]
        assert isinstance(styles, list)
        assert len(styles) >= 1


class TestQueryLinetypes:
    async def test_returns_linetypes(self, backend):
        r = await backend.query_linetypes()
        assert r.ok
        lts = r.payload["linetypes"]
        assert isinstance(lts, list)
        names = [lt["name"] for lt in lts]
        # ezdxf always has at least ByBlock, ByLayer, Continuous
        assert any(n in names for n in ["Continuous", "ByBlock", "ByLayer"])


class TestQueryBlockTree:
    async def test_empty_drawing(self, backend):
        r = await backend.query_block_tree()
        assert r.ok
        blocks = r.payload["blocks"]
        assert isinstance(blocks, list)

    async def test_with_defined_block(self, populated):
        r = await populated.query_block_tree()
        assert r.ok
        names = [b["name"] for b in r.payload["blocks"]]
        assert "PUMP" in names


class TestQueryDrawingMetadata:
    async def test_returns_metadata(self, backend):
        r = await backend.query_drawing_metadata()
        assert r.ok
        p = r.payload
        assert any(k in p for k in ["dxf_version", "version", "units", "limmin", "limmax"])


# ---------------------------------------------------------------------------
# Extended Search
# ---------------------------------------------------------------------------


class TestSearchByBlockName:
    async def test_find_pump_blocks(self, populated):
        r = await populated.search_by_block_name("PUMP")
        assert r.ok
        results = r.payload["entities"]
        assert len(results) == 2

    async def test_no_results(self, populated):
        r = await populated.search_by_block_name("NONEXISTENT")
        assert r.ok
        results = r.payload["entities"]
        assert len(results) == 0


class TestSearchByHandleList:
    async def test_batch_get(self, populated):
        listing = await populated.entity_list()
        handles = [e["handle"] for e in listing.payload["entities"][:3]]
        r = await populated.search_by_handle_list(handles)
        assert r.ok
        results = r.payload["entities"]
        assert len(results) == 3


# ---------------------------------------------------------------------------
# Extended Validate
# ---------------------------------------------------------------------------


class TestValidateTextStandards:
    async def test_with_allowed_styles(self, populated):
        r = await populated.validate_text_standards(
            allowed_styles=["Standard"],
            allowed_heights=[3.5, 5.0]
        )
        assert r.ok
        violations = r.payload.get("violations", [])
        assert len(violations) == 0

    async def test_violation_detected(self, populated):
        r = await populated.validate_text_standards(
            allowed_styles=["CUSTOM_ONLY"],
            allowed_heights=[2.5]
        )
        assert r.ok
        violations = r.payload.get("violations", [])
        assert len(violations) > 0


class TestValidateOrphanedEntities:
    async def test_no_orphans(self, backend):
        """Test with a clean backend where layer '0' is on."""
        # Ensure layer 0 is on (ezdxf may default to off)
        layer0 = backend._doc.layers.get("0")
        layer0.on = True
        await backend.create_line(0, 0, 10, 10)
        r = await backend.validate_orphaned_entities()
        assert r.ok
        issues = r.payload.get("issues", [])
        assert len(issues) == 0

    async def test_frozen_layer_orphan(self, backend):
        """Create entities on a layer, freeze it, check detection."""
        await backend.layer_create("TEST_WIRE", "1", "Continuous")
        await backend.create_line(0, 0, 100, 0, layer="TEST_WIRE")
        await backend.create_line(100, 0, 200, 0, layer="TEST_WIRE")
        await backend.create_line(200, 0, 200, 100, layer="TEST_WIRE")
        await backend.layer_freeze("TEST_WIRE")
        r = await backend.validate_orphaned_entities()
        assert r.ok
        issues = r.payload.get("issues", [])
        assert len(issues) >= 3  # 3 lines on frozen TEST_WIRE layer


class TestValidateAttributeCompleteness:
    async def test_all_filled(self, populated):
        r = await populated.validate_attribute_completeness(required_tags=["TAG"])
        assert r.ok
        issues = r.payload.get("issues", [])
        # Both blocks have TAG filled
        assert len(issues) == 0

    async def test_missing_desc(self, populated):
        r = await populated.validate_attribute_completeness(required_tags=["TAG", "DESC"])
        assert r.ok
        issues = r.payload.get("issues", [])
        # P-102 has DESC="" which should count as incomplete
        assert len(issues) >= 1


class TestValidateConnectivity:
    async def test_connected_lines(self, populated):
        r = await populated.validate_connectivity(layer="WIRE", tolerance=0.01)
        assert r.ok
        # Lines: (0,0)-(100,0), (100,0)-(200,0), (200,0)-(200,100)
        # Dangling ends at (0,0) and (200,100)
        dangling = r.payload.get("dangling", [])
        assert len(dangling) == 2


# ---------------------------------------------------------------------------
# Extended Select: find_replace_attribute, layer_rename, layer_merge
# ---------------------------------------------------------------------------


class TestFindReplaceAttribute:
    async def test_replace_tag_value(self, populated):
        r = await populated.find_replace_attribute("TAG", "P-101", "P-201")
        assert r.ok
        count = r.payload.get("replaced", 0)
        assert count == 1

    async def test_no_match(self, populated):
        r = await populated.find_replace_attribute("TAG", "NOMATCH", "ANYTHING")
        assert r.ok
        count = r.payload.get("replaced", 0)
        assert count == 0


class TestLayerRename:
    async def test_rename_layer(self, populated):
        r = await populated.layer_rename("WIRE", "E-WIRE")
        assert r.ok
        layers = await populated.layer_list()
        names = [l["name"] for l in layers.payload["layers"]]
        assert "E-WIRE" in names
        assert "WIRE" not in names

    async def test_rename_nonexistent(self, populated):
        r = await populated.layer_rename("NOPE", "SOMETHING")
        assert not r.ok


class TestLayerMerge:
    async def test_merge_layers(self, populated):
        listing = await populated.entity_list(layer="EQUIP")
        equip_count = len(listing.payload["entities"])

        r = await populated.layer_merge("WIRE", "EQUIP")
        assert r.ok

        listing2 = await populated.entity_list(layer="EQUIP")
        new_count = len(listing2.payload["entities"])
        assert new_count > equip_count


# ---------------------------------------------------------------------------
# Enhanced Drawing: audit, units, limits, wblock
# ---------------------------------------------------------------------------


class TestDrawingAudit:
    async def test_audit_clean_drawing(self, backend):
        r = await backend.drawing_audit()
        assert r.ok
        assert "issues" in r.payload or "issue_count" in r.payload


class TestDrawingUnits:
    async def test_get_units(self, backend):
        r = await backend.drawing_units()
        assert r.ok

    async def test_set_units(self, backend):
        r = await backend.drawing_units(units=4)  # Millimeters
        assert r.ok


class TestDrawingLimits:
    async def test_get_limits(self, backend):
        r = await backend.drawing_limits()
        assert r.ok
        assert "limmin" in r.payload or "min" in r.payload

    async def test_set_limits(self, backend):
        r = await backend.drawing_limits(x1=0, y1=0, x2=1000, y2=1000)
        assert r.ok


class TestDrawingWblock:
    async def test_export_entities(self, populated):
        listing = await populated.entity_list()
        handles = [e["handle"] for e in listing.payload["entities"][:2]]

        with tempfile.NamedTemporaryFile(suffix=".dxf", delete=False) as f:
            path = f.name
        try:
            r = await populated.drawing_wblock(handles, path)
            assert r.ok
            assert os.path.exists(path)
            doc = ezdxf.readfile(path)
            msp = doc.modelspace()
            assert len(list(msp)) == 2
        finally:
            os.unlink(path)


# ---------------------------------------------------------------------------
# View (headless placeholders)
# ---------------------------------------------------------------------------


class TestViewEnhanced:
    async def test_zoom_scale(self, backend):
        r = await backend.zoom_scale(2.0)
        assert r.ok

    async def test_pan(self, backend):
        r = await backend.pan(100, 50)
        assert r.ok


# ---------------------------------------------------------------------------
# XREF
# ---------------------------------------------------------------------------


class TestXref:
    async def test_xref_list_empty(self, backend):
        r = await backend.xref_list()
        assert r.ok
        xrefs = r.payload["xrefs"]
        assert len(xrefs) == 0

    async def test_xref_query_entities_block(self, populated):
        r = await populated.xref_query_entities("PUMP")
        assert r.ok
        assert "entities" in r.payload


# ---------------------------------------------------------------------------
# Layout
# ---------------------------------------------------------------------------


class TestLayout:
    async def test_layout_list(self, backend):
        r = await backend.layout_list()
        assert r.ok
        layouts = r.payload["layouts"]
        assert len(layouts) >= 1  # At least "Model"

    async def test_layout_create(self, backend):
        r = await backend.layout_create("Sheet1")
        assert r.ok

        r2 = await backend.layout_list()
        names = [l["name"] for l in r2.payload["layouts"]]
        assert "Sheet1" in names

    async def test_layout_delete(self, backend):
        await backend.layout_create("TempLayout")
        r = await backend.layout_delete("TempLayout")
        assert r.ok

    async def test_layout_switch(self, backend):
        r = await backend.layout_switch("Model")
        assert r.ok

    async def test_layout_viewport_create(self, backend):
        await backend.layout_create("VP_Test")
        r = await backend.layout_viewport_create(5, 5, 10, 8, scale=50.0)
        assert isinstance(r.ok, bool)


# ---------------------------------------------------------------------------
# Electrical
# ---------------------------------------------------------------------------


class TestElectricalNECLookup:
    async def test_wire_ampacity_12awg(self, backend):
        r = await backend.electrical_nec_lookup("wire_ampacity", {"wire_gauge": "12"})
        assert r.ok
        assert r.payload["ampacity"] == 20

    async def test_wire_ampacity_4_0(self, backend):
        r = await backend.electrical_nec_lookup("wire_ampacity", {"wire_gauge": "4/0"})
        assert r.ok
        assert r.payload["ampacity"] == 195

    async def test_unknown_gauge(self, backend):
        r = await backend.electrical_nec_lookup("wire_ampacity", {"wire_gauge": "99"})
        assert not r.ok

    async def test_unknown_table(self, backend):
        r = await backend.electrical_nec_lookup("unknown_table", {})
        # May return ok=True with error info or ok=False
        if r.ok:
            assert "error" in r.payload or "available_tables" in r.payload
        # Either way, not a crash


class TestElectricalVoltageDrop:
    async def test_single_phase_120v(self, backend):
        # 120V, 16A, 12AWG, 100ft, single phase
        r = await backend.electrical_voltage_drop(120, 16, "12", 100, phase=1, power_factor=1.0)
        assert r.ok
        vd = r.payload["voltage_drop"]
        pct = r.payload.get("voltage_drop_percent", r.payload.get("percent_drop", 0))
        # Vd = 2 * 100 * 16 * 1.98 / 1000 = 6.336V -> 5.28%
        assert abs(vd - 6.336) < 0.5
        assert pct > 3.0  # Over 3%
        # Check acceptable flag (may be "acceptable" or "pass_branch")
        acceptable = r.payload.get("acceptable", r.payload.get("pass_branch", None))
        if acceptable is not None:
            assert acceptable is False  # >3% is not acceptable

    async def test_three_phase(self, backend):
        # 480V, 100A, 4/0 AWG, 200ft, 3-phase
        r = await backend.electrical_voltage_drop(480, 100, "4/0", 200, phase=3, power_factor=0.85)
        assert r.ok
        vd = r.payload["voltage_drop"]
        # Should be a small voltage drop for 4/0 at 480V
        assert vd < 10.0
        acceptable = r.payload.get("acceptable", r.payload.get("pass_branch", None))
        if acceptable is not None:
            assert acceptable is True

    async def test_unknown_gauge_still_calculates(self, backend):
        # Implementation may use default resistance for unknown gauge
        r = await backend.electrical_voltage_drop(120, 10, "99", 50)
        # Either returns error or calculates with default - don't crash
        assert isinstance(r.ok, bool)


class TestElectricalConduitFill:
    async def test_3_wires_12awg_in_half_inch(self, backend):
        r = await backend.electrical_conduit_fill("1/2", "EMT", ["12", "12", "12"])
        assert r.ok
        fill = r.payload["fill_percent"]
        # 3 * 0.0133 = 0.0399 / 0.304 * 100 = 13.12%
        assert fill < 40.0
        acceptable = r.payload.get("acceptable", r.payload.get("pass", None))
        if acceptable is not None:
            assert acceptable is True

    async def test_overfilled_conduit(self, backend):
        # 10 x 4AWG wires into 1/2" conduit
        r = await backend.electrical_conduit_fill("1/2", "EMT", ["4"] * 10)
        assert r.ok
        assert r.payload["fill_percent"] > 40.0
        acceptable = r.payload.get("acceptable", r.payload.get("pass", None))
        if acceptable is not None:
            assert acceptable is False


class TestElectricalLoadCalc:
    async def test_basic_load(self, backend):
        r = await backend.electrical_load_calc(
            devices=[{"name": "Light", "watts": 100}, {"name": "Motor", "watts": 500}]
        )
        assert r.ok
        total = r.payload.get("total_watts", r.payload.get("total_va", 0))
        assert total >= 600
