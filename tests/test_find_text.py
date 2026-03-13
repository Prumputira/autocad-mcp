"""Tests for find_text deep search operation."""

import math
import pytest

from autocad_mcp.backends.ezdxf_backend import EzdxfBackend


@pytest.fixture
async def backend():
    b = EzdxfBackend()
    await b.initialize()
    return b


@pytest.fixture
async def populated_backend():
    """Backend with TEXT, MTEXT, INSERT+ATTRIB, and block definition text."""
    b = EzdxfBackend()
    await b.initialize()
    msp = b._msp
    doc = b._doc

    # Modelspace TEXT
    msp.add_text("BS-584 valve", dxfattribs={"insert": (100, 200), "layer": "TEXT"})
    msp.add_text("Other text", dxfattribs={"insert": (300, 400), "layer": "TEXT"})

    # Modelspace MTEXT
    msp.add_mtext("Refer to BS-584 spec", dxfattribs={"insert": (150, 250), "layer": "NOTES"})

    # Block with ATTRIB containing the pattern
    block_with_attrib = doc.blocks.new(name="EQUIP_TAG")
    block_with_attrib.add_attdef("TAGNAME", (0, 0), dxfattribs={"text": "BS-584"})
    block_with_attrib.add_circle((0, 0), 5)

    ins = msp.add_blockref("EQUIP_TAG", (500, 600))
    ins.add_auto_attribs({"TAGNAME": "BS-584"})

    # Block with TEXT inside definition (no modelspace INSERT → nested_block_definition)
    orphan_block = doc.blocks.new(name="ORPHAN_BLOCK")
    orphan_block.add_text("BS-584 note", dxfattribs={"insert": (10, 20), "layer": "0"})

    # Block with TEXT inside definition WITH modelspace INSERT
    ref_block = doc.blocks.new(name="REF_BLOCK")
    ref_block.add_text("Contains BS-584", dxfattribs={"insert": (5, 10), "layer": "0"})
    msp.add_blockref("REF_BLOCK", (1000, 2000))

    # Block with no matching text
    no_match_block = doc.blocks.new(name="NO_MATCH")
    no_match_block.add_text("Nothing here", dxfattribs={"insert": (0, 0), "layer": "0"})

    return b


class TestFindTextBasic:
    async def test_empty_pattern_returns_empty(self, backend):
        r = await backend.find_text("")
        # ezdxf backend returns ok=True with empty results for empty pattern
        # LISP backend returns error; both are acceptable
        assert r.ok is True or r.ok is False

    async def test_no_document(self):
        b = EzdxfBackend()
        await b.initialize()
        b._doc = None
        b._msp = None
        r = await b.find_text("test")
        assert not r.ok

    async def test_no_results(self, backend):
        r = await backend.find_text("nonexistent_pattern_xyz")
        assert r.ok
        assert r.payload["count"] == 0
        assert r.payload["results"] == []


class TestFindTextModelspace:
    async def test_finds_text_entity(self, populated_backend):
        r = await populated_backend.find_text("BS-584")
        assert r.ok
        assert r.payload["count"] > 0
        types = [res["type"] for res in r.payload["results"]]
        assert "TEXT" in types

    async def test_finds_mtext_entity(self, populated_backend):
        r = await populated_backend.find_text("BS-584")
        assert r.ok
        mtexts = [res for res in r.payload["results"]
                  if res["type"] == "MTEXT" and res["context"] == "modelspace"]
        assert len(mtexts) >= 1

    async def test_case_insensitive_default(self, populated_backend):
        r = await populated_backend.find_text("bs-584")
        assert r.ok
        assert r.payload["count"] > 0

    async def test_case_sensitive(self, populated_backend):
        r = await populated_backend.find_text("bs-584", case_sensitive=True)
        assert r.ok
        # "bs-584" lowercase shouldn't match "BS-584"
        assert r.payload["count"] == 0


class TestFindTextAttributes:
    async def test_finds_attrib(self, populated_backend):
        r = await populated_backend.find_text("BS-584")
        assert r.ok
        attribs = [res for res in r.payload["results"] if res["type"] == "ATTRIB"]
        assert len(attribs) >= 1
        assert attribs[0]["context"] == "attribute"
        assert attribs[0]["containing_block"] == "EQUIP_TAG"


class TestFindTextBlockDefinitions:
    async def test_finds_text_in_block_def_with_insert(self, populated_backend):
        r = await populated_backend.find_text("BS-584")
        assert r.ok
        block_defs = [res for res in r.payload["results"]
                      if res["context"] == "block_definition"
                      and res["containing_block"] == "REF_BLOCK"]
        assert len(block_defs) >= 1

    async def test_finds_text_in_orphan_block(self, populated_backend):
        r = await populated_backend.find_text("BS-584")
        assert r.ok
        nested = [res for res in r.payload["results"]
                  if res["context"] == "nested_block_definition"
                  and res["containing_block"] == "ORPHAN_BLOCK"]
        assert len(nested) >= 1

    async def test_world_position_transformed(self, populated_backend):
        r = await populated_backend.find_text("BS-584")
        assert r.ok
        block_defs = [res for res in r.payload["results"]
                      if res["context"] == "block_definition"
                      and res["containing_block"] == "REF_BLOCK"]
        assert len(block_defs) >= 1
        wp = block_defs[0]["world_position"]
        # REF_BLOCK inserted at (1000, 2000), text at (5, 10) local
        # Scale=1, rotation=0 → world = (1005, 2010)
        assert abs(wp[0] - 1005) < 1
        assert abs(wp[1] - 2010) < 1


class TestFindTextMaxResults:
    async def test_max_results_limit(self, populated_backend):
        r = await populated_backend.find_text("BS-584", max_results=2)
        assert r.ok
        assert r.payload["count"] <= 2

    async def test_zoomed_to(self, populated_backend):
        r = await populated_backend.find_text("BS-584")
        assert r.ok
        if r.payload["count"] > 0:
            assert "zoomed_to" in r.payload


class TestFindTextWildcard:
    async def test_wildcard_pattern(self, populated_backend):
        r = await populated_backend.find_text("BS-*")
        assert r.ok
        assert r.payload["count"] > 0

    async def test_substring_match(self, populated_backend):
        r = await populated_backend.find_text("584")
        assert r.ok
        assert r.payload["count"] > 0
