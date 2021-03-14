# encoding: utf-8

"""lxml custom element class for CT_GraphicalObjectFrame XML element."""

from pptx.oxml import parse_xml
from pptx.oxml.chart.chart import CT_Chart
from pptx.oxml.ns import nsdecls
from pptx.oxml.shapes.shared import BaseShapeElement
from pptx.oxml.simpletypes import XsdString
from pptx.oxml.table import CT_Table
from pptx.oxml.xmlchemy import (
    BaseOxmlElement,
    OneAndOnlyOne,
    OptionalAttribute,
    RequiredAttribute,
    ZeroOrOne,
)
from pptx.spec import (
    GRAPHIC_DATA_URI_CHART,
    GRAPHIC_DATA_URI_OLEOBJ,
    GRAPHIC_DATA_URI_TABLE,
)


class CT_GraphicalObject(BaseOxmlElement):
    """
    ``<a:graphic>`` element, which is the container for the reference to or
    definition of the framed graphical object (table, chart, etc.).
    """

    graphicData = OneAndOnlyOne("a:graphicData")

    @property
    def chart(self):
        """
        The ``<c:chart>`` grandchild element, or |None| if not present.
        """
        return self.graphicData.chart


class CT_GraphicalObjectData(BaseShapeElement):
    """
    ``<p:graphicData>`` element, the direct container for a table, a chart,
    or another graphical object.
    """

    chart = ZeroOrOne("c:chart")
    tbl = ZeroOrOne("a:tbl")
    uri = RequiredAttribute("uri", XsdString)

    @property
    def blob_rId(self):
        """Optional "r:id" attribute value of `<p:oleObj>` descendent element.

        This value is `None` when this `p:graphicData` element does not enclose an OLE
        object. This value could also be `None` if an enclosed OLE object does not
        specify this attribute (it is specified optional in the schema) but so far, all
        OLE objects we've encountered specify this value.
        """
        return None if self._oleObj is None else self._oleObj.rId

    @property
    def is_embedded_ole_obj(self):
        """Optional boolean indicating an embedded OLE object.

        Returns `None` when this `p:graphicData` element does not enclose an OLE object.
        `True` indicates an embedded OLE object and `False` indicates a linked OLE
        object.
        """
        return None if self._oleObj is None else self._oleObj.is_embedded

    @property
    def _oleObj(self):
        """Optional `<p:oleObj>` element contained in this `p:graphicData' element.

        Returns `None` when this graphic-data element does not enclose an OLE object.
        Note that this returns the last `p:oleObj` element found. There can be more
        than one `p:oleObj` element because an `<mc.AlternateContent>` element may
        appear as the child of `p:graphicData` and that alternate-content subtree can
        contain multiple compatibility choices. The last one should suit best for
        reading purposes because it contains the lowest common denominator.
        """
        oleObjs = self.xpath(".//p:oleObj")
        return oleObjs[-1] if oleObjs else None


class CT_GraphicalObjectFrame(BaseShapeElement):
    """
    ``<p:graphicFrame>`` element, which is a container for a table, a chart,
    or another graphical object.
    """

    nvGraphicFramePr = OneAndOnlyOne("p:nvGraphicFramePr")
    xfrm = OneAndOnlyOne("p:xfrm")
    graphic = OneAndOnlyOne("a:graphic")

    @property
    def chart(self):
        """
        The ``<c:chart>`` great-grandchild element, or |None| if not present.
        """
        return self.graphic.chart

    @property
    def chart_rId(self):
        """
        The ``rId`` attribute of the ``<c:chart>`` great-grandchild element,
        or |None| if not present.
        """
        chart = self.chart
        if chart is None:
            return None
        return chart.rId

    def get_or_add_xfrm(self):
        """
        Return the required ``<p:xfrm>`` child element. Overrides version on
        BaseShapeElement.
        """
        return self.xfrm

    @property
    def graphicData(self):
        """`<a:graphicData> grandchild of this graphic-frame element."""
        return self.graphic.graphicData

    @property
    def graphicData_uri(self):
        """str value of `uri` attribute of `<a:graphicData> grandchild."""
        return self.graphic.graphicData.uri

    @property
    def has_oleobj(self):
        """True for graphicFrame containing an OLE object, False otherwise."""
        return self.graphicData.uri == GRAPHIC_DATA_URI_OLEOBJ

    @property
    def is_embedded_ole_obj(self):
        """Optional boolean indicating an embedded OLE object.

        Returns `None` when this `p:graphicFrame` element does not enclose an OLE
        object. `True` indicates an embedded OLE object and `False` indicates a linked
        OLE object.
        """
        return self.graphicData.is_embedded_ole_obj

    @classmethod
    def new_chart_graphicFrame(cls, id_, name, rId, x, y, cx, cy):
        """
        Return a ``<p:graphicFrame>`` element tree populated with a chart
        element.
        """
        graphicFrame = CT_GraphicalObjectFrame.new_graphicFrame(id_, name, x, y, cx, cy)
        graphicData = graphicFrame.graphic.graphicData
        graphicData.uri = GRAPHIC_DATA_URI_CHART
        graphicData.append(CT_Chart.new_chart(rId))
        return graphicFrame

    @classmethod
    def new_graphicFrame(cls, id_, name, x, y, cx, cy):
        """
        Return a new ``<p:graphicFrame>`` element tree suitable for
        containing a table or chart. Note that a graphicFrame element is not
        a valid shape until it contains a graphical object such as a table.
        """
        xml = cls._graphicFrame_tmpl() % (id_, name, x, y, cx, cy)
        graphicFrame = parse_xml(xml)
        return graphicFrame

    @classmethod
    def new_table_graphicFrame(cls, id_, name, rows, cols, x, y, cx, cy):
        """
        Return a ``<p:graphicFrame>`` element tree populated with a table
        element.
        """
        graphicFrame = cls.new_graphicFrame(id_, name, x, y, cx, cy)
        graphicFrame.graphic.graphicData.uri = GRAPHIC_DATA_URI_TABLE
        graphicFrame.graphic.graphicData.append(CT_Table.new_tbl(rows, cols, cx, cy))
        return graphicFrame

    @classmethod
    def _graphicFrame_tmpl(cls):
        return (
            "<p:graphicFrame %s>\n"
            "  <p:nvGraphicFramePr>\n"
            '    <p:cNvPr id="%s" name="%s"/>\n'
            "    <p:cNvGraphicFramePr>\n"
            '      <a:graphicFrameLocks noGrp="1"/>\n'
            "    </p:cNvGraphicFramePr>\n"
            "    <p:nvPr/>\n"
            "  </p:nvGraphicFramePr>\n"
            "  <p:xfrm>\n"
            '    <a:off x="%s" y="%s"/>\n'
            '    <a:ext cx="%s" cy="%s"/>\n'
            "  </p:xfrm>\n"
            "  <a:graphic>\n"
            "    <a:graphicData/>\n"
            "  </a:graphic>\n"
            "</p:graphicFrame>"
            % (nsdecls("a", "p"), "%d", "%s", "%d", "%d", "%d", "%d")
        )


class CT_GraphicalObjectFrameNonVisual(BaseOxmlElement):
    """`<p:nvGraphicFramePr>` element.

    This contains the non-visual properties of a graphic frame, such as name, id, etc.
    """

    cNvPr = OneAndOnlyOne("p:cNvPr")
    nvPr = OneAndOnlyOne("p:nvPr")


class CT_OleObject(BaseOxmlElement):
    """`<p:oleObj>` element, container for an OLE object (e.g. Excel file).

    An OLE object can be either linked or embedded (hence the name).
    """

    rId = OptionalAttribute("r:id", XsdString)

    @property
    def is_embedded(self):
        """True when this OLE object is embedded, False when it is linked."""
        return True if len(self.xpath("./p:embed")) > 0 else False
