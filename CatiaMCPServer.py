#2026.4.21
import win32com.client
import pywintypes #这是.pyd
from typing import Any #暂时未用
import httpx #暂时未用
from mcp.server.fastmcp import FastMCP
import sys
import logging #暂时未用
import math
import re


mcp = FastMCP("Catia") #初始化FastMCP Server

#Constants 配置常量暂无
#API_BASE= None
#UESR_AGENT=None

#全局常量 2026.4.24
catia=None
try:   
    catia = win32com.client.Dispatch("Catia.Application")
    print("get the IDispatch-catia", file=sys.stderr)
except pywintypes.com_error:
    print("Don't have Catia or IDispatch is failed.", file=sys.stderr)


#Helper functions 辅助函数暂无

#mcptool
@mcp.tool()#2026.4.30, 2026.5.1
async def open_catia(user_order:bool)->str:
    '''open the catia application

        Args:
        user_order: True or False; True means the user clearly asks AI to use catia, otherwise,it is False.
    '''
    global catia
    if user_order==True:
        try:
            catia.Visible=True
        except:
            return "can't open catia"
        else:
            visibility=catia.Visible
            if visibility==True:
                return "open the catia"
            else:
                return "no sure"
    else:
        return "wait for the user's ask"

@mcp.tool()#2026.4.21，2026.4.30, 2026.6.4
async def add_partdoc(partdoc_name:str,saved:bool)->str:
    '''create a Part Document

    Args:
        partdoc_name: The Name of the Partdocument
        #未使用；后面保存文件代替了这个传入
        saved: the action to save the other documents; True means the user have saved the others documents, otherwise, it is False.
    '''
    global catia
    try:
        docs_number=catia.Documents.Count
    except pywintypes.com_error:
        return "IDispatch is failed"
    else:
        if docs_number!=0 and saved!=True:
            return f"There are {docs_number} Documents, that are already opend. Please check them, and save or close them"
        else:
            try:
                newpartdoc=catia.Documents.Add("Part")
            except pywintypes.com_error:
                return "IDispatch is failed"
            else:
                try:
                    activedoc=catia.ActiveDocument
                    return f"The PartDocument{newpartdoc.Name} is already created and opened. Don't forget to save it"
                except:
                    catia.quit()
                    return "The root for Claude for Desktop and Catia are not in same level."
                    #进一步检查Claude for Desktop 和 Catia 为同一级权限

@mcp.tool()#2026.6.4
async def add_drawingdoc(drawingdoc_name:str,saved:bool)->str:
    '''create a Drawing Document

    Args:
        drawingdoc_name: The Name of the Drawingdocument
        #未使用；后面保存文件代替了这个传入
        saved: the action to save the other documents; True means the user have saved the others documents, otherwise, it is False.
    '''
    global catia
    try:
        docs_number=catia.Documents.Count
    except pywintypes.com_error:
        return "IDispatch is failed"
    else:
        if docs_number!=0 and saved!=True:
            return f"There are {docs_number} Documents, that are already opend. Please check them, and save or close them"
        else:
            try:
                newdrawingdoc=catia.Documents.Add("Drawing")
            except pywintypes.com_error:
                return "IDispatch is failed"
            else:
                try:
                    activedoc=catia.ActiveDocument
                    return f"The DrawingDocument{newdrawingdoc.Name} is already created and opened. Don't forget to save it"
                except:
                    catia.quit()
                    return "The root for Claude for Desktop and Catia are not in same level."
                    #进一步检查Claude for Desktop 和 Catia 为同一级权限            


@mcp.tool()#2026.4.21,2026.5.1
async def saveas_partdoc(partdoc_name:str, partdoc_path:str)->str:
    ''' save a Part document with a new name and path

        Args:
            partdoc_name: The name for the Part document
            partdoc_path: The path for the Part document
    '''
    global catia
    #检查CATDocView和CATTEMP环境变量-文件要求（Use case），可优化
    try:
        doc_path=catia.SystemService.Environ("CATDocView")
        docpath_exist=catia.FileSystem.FolderExists(doc_path)
        print(f"{doc_path}", file=sys.stderr)
    except:
        return "CATDocView is not right"
    try:
        tmp_path=catia.SystemService.Environ("CATTemp")
        tmppath_exist=catia.FileSystem.FolderExists(tmp_path)
        print(f"{tmp_path}", file=sys.stderr)
    except:
        return "CATTEMP is not right"
    try:
        activedoc=catia.ActiveDocument
        activedoc.SaveAS(f"{partdoc_path}\{partdoc_name}.CATPart")
        return(f"{partdoc_name} save in {partdoc_path}")
    except:
        return("can't save this Document")

@mcp.tool()#2026.6.4
async def saveas_drawingdoc(drawingdoc_name:str, drawingdoc_path:str)->str:
    ''' save a Drawing document with a new name and path

        Args:
            drawingdoc_name: The name for the drawing document
            drawingdoc_path: The path for the drawing document
    '''
    global catia
    #检查CATDocView和CATTEMP环境变量-文件要求（Use case），可优化
    try:
        doc_path=catia.SystemService.Environ("CATDocView")
        docpath_exist=catia.FileSystem.FolderExists(doc_path)
        print(f"{doc_path}", file=sys.stderr)
    except:
        return "CATDocView is not right"
    try:
        tmp_path=catia.SystemService.Environ("CATTemp")
        tmppath_exist=catia.FileSystem.FolderExists(tmp_path)
        print(f"{tmp_path}", file=sys.stderr)
    except:
        return "CATTEMP is not right"
    try:
        activedoc=catia.ActiveDocument
        activedoc.SaveAS(f"{drawingdoc_path}\{drawingdoc_name}.CATDrawing")
        return(f"{drawingdoc_name} save in {drawingdoc_path}")
    except:
        return("can't save this Document")


@mcp.tool()#2026.4.22, 2026.6.4
async def open_doc(docpath:str)->str:
    '''open a document, when user ask

        Args:
        doc_path: The path for the requested Document
    '''
    global catia
    try:
        catia.Documents.open(docpath)
        return "The Document is opend."
    except:
        return "Can't open the Document."

@mcp.tool()#2026.5.1, 2026.6.4
async def save_doc(saveddoc_name:str)->str:
    '''save a document, which isn't just created or which has been saved

        Args:
        saveddoc_name: the name of the document
    '''
    global catia
    try:
        doc_path=catia.SystemService.Environ("CATDocView")
        docpath_exist=catia.FileSystem.FolderExists(doc_path)
        print(f"{doc_path}", file=sys.stderr)
    except:
        return "CATDocView is not right"
        #待优化
    else:
        mydoc=catia.Documents.Item(saveddoc_name)
        mydoc.Save()
        return f"{mydoc.Name} has been saved"

@mcp.tool()#2026.4.30-2026.5.1
async def close_doc(doc_name:str,use_order:bool)->str:
    '''close a document

        Args:
        doc_name: the name of the documents,that user want to close
        use_order: The user clearly asks to close the document; True means the user knows, that the document has been changed, and still asks to close, otherwise, it is False.
    '''
    global catia
    mydoc=catia.Documents.Item(doc_name)
    haschanged=mydoc.Saved
    #AI会不会先询问保存，和人一样带有不确定性，所以设置以下步骤，即使有save_partdoc的tool
    if haschanged==True:
        mydoc.Close()
        return "the Document closed"
    else:
        if use_order==True:
            try:
                doc_path=catia.SystemService.Environ("CATDocView")
                docpath_exist=catia.FileSystem.FolderExists(doc_path)
                print(f"{doc_path}", file=sys.stderr)
            except:
                return "CATDocView is not right"
            try:
                tmp_path=catia.SystemService.Environ("CATTemp")
                tmppath_exist=catia.FileSystem.FolderExists(tmp_path)
                print(f"{tmp_path}", file=sys.stderr)
            except:
                return "CATTEMP is not right"
            #以上步骤可在后续开发中优化，变为服务器本身状态的监控，save_partdoc也有
            doc_path=mydoc.Path
            if doc_path:
                mydoc.Save()
                mydoc.Close()
                return f"The document{mydoc.Name} is closed and in {doc_path}"
            else:
                return "It's a new Document, please save it with new path and name"
        else:
            return "the document has been change, wait for the user's order"


@mcp.tool()#2026.5.2
#添加这个工具仅仅是个人习惯，再原点（0，0，0）使用坐标轴，其本身的功能远比这个tool强大和复杂
async def add_axissystemstandard()->str:
    '''create a  standard axissystem
    '''
    global catia
    rootpart=catia.ActiveDocument.Part
    newaxissystem=rootpart.AxisSystems.Add()
    newaxissystem.OriginType=1
    newaxissystem.Type=0
    newaxissystem.IsCurrent=1
    rootpart.Update()
    return "create a standard axissystem"

@mcp.tool()#2026.4.24-2026.4.25,2026.5.2,2026.5.24
async def add_sketch(body_name:str,plane_name:str)->str:
    '''create a sketch form origin plane 

        Args:
        body_name: the name of the body
        plane_name: choosed plane name for the sketch
    '''
    global catia
    rootpart=catia.ActiveDocument.Part

    if plane_name=="PlaneXY":
        plane=rootpart.OriginElements.PlaneXY
    elif plane_name=="PlaneYZ":
        plane=rootpart.OriginElements.PlaneYZ
    elif plane_name=="PlaneZX":
        plane=rootpart.OriginElements.PlaneZX
    else:
        return "planename isn't from origin plane"
    mybody=rootpart.Bodies.Item(body_name)
    try:
        newsketch=mybody.Sketches.Add(plane)
        return f"added a sketch: {newsketch.Name}"
    except:
        return "can't added a sketch"

@mcp.tool()#2026.4.28,2026.5.24
async def create_point2d(body_name:str,sketch_name:str,x:float,y:float)->str:
    '''create a point

        Args:
        body_name: the name of the body
        sketch_name: the name of the sketch
        x: The X coordinate of point 
        y: The Y coordinate of point   
    '''
    global catia
    rootpart=catia.ActiveDocument.Part
    mybody=rootpart.Bodies.Item(body_name)
    try:
        mysketch=mybody.Sketches.Item(sketch_name)
    except:
        return f"can't find the sketch({sketch_name}) in body({body_name})"
    try:
        factory2d=mysketch.OpenEdition()
        mypoint=factory2d.CreatePoint(x,y)
        factory2d=mysketch.CloseEdition()
        rootpart.Update()
        return f"Control Point {mypoint.Name} ({x},{y})"
    except:
        factory2d=mysketch.CloseEdition()
        rootpart.Update()
        return "can't create a point"

@mcp.tool()#2026.4.25,2026.4.28,2026.5.24
async def create_closedcircle2d(body_name:str,sketch_name:str,x:float, y:float, r:float)->str:
    '''create a 2D closed circle with center point in sketch

        Args:
        body_name: the name of the body
        sketch_name: the name of the sketch
        x: Horizontal axis coordinates form same  point
        y: Longitudinal axis coordinates form same point
        r: Radius
    '''
    global catia
    rootpart=catia.ActiveDocument.Part
    mybody=rootpart.Bodies.Item(body_name)
    try:
        mysketch=mybody.Sketches.Item(sketch_name)
    except:
        return f"can't find the sketch({sketch_name}) in body({body_name})"
    try:    
        factory2d=mysketch.OpenEdition()
        centerpoint=factory2d.CreatePoint(x,y)
        circle2d=factory2d.CreateClosedCircle(x,y,r)
        try:
            circle2d.CenterPoint=centerpoint
        except:
            factory2d=mysketch.CloseEdition()
            rootpart.Update()
            return "no center punkt"

        factory2d=mysketch.CloseEdition()
        rootpart.Update()
        return f"create a Closedcircle {circle2d.Name},({x},{y},{r}) in {mysketch.Name}"
    except:
        factory2d=mysketch.CloseEdition()
        rootpart.Update()
        return "can't create a closed 2D circle"

@mcp.tool()#2026.4.26,2026.4.28,2026.5.24
async def create_circle2d(body_name:str,sketch_name:str,x:float,y:float,r:float,startparam:float,endparam:float)->str:
    '''create a 2d circle in sketch

        Args:
        body_name: the name of the body
        sketch_name: the name of the sketch
        x:The X coordinate of the circle center
        y:The Y coordinate of the circle center
        r:The radius of the circle
        startparam:The beginning parameter of the circle.This parameter is an angle value between 0 included and 2PI excluded. Parameter values are computed from the axis horizontal direction in the trigonometrical direction.
        endparam:The end parameter of the circle.This parameter may take any value between iStartParam excluded and 4PI included.
    '''
    global catia
    rootpart=catia.ActiveDocument.Part
    mybody=rootpart.Bodies.Item(body_name)
    try:
        mysketch=mybody.Sketches.Item(sketch_name)
    except:
        return f"can't find the sketch({sketch_name}) in body({body_name})"
    try:    
        factory2d=mysketch.OpenEdition()
        centerpoint=factory2d.CreatePoint(x,y)
        circle2d=factory2d.CreateCircle(x,y,r,startparam,endparam)
        circle2d.CenterPoint=centerpoint
        try:
            x1=x+math.cos(startparam)
            y1=y+math.sin(startparam)
            startpoint=factory2d.CreatePoint(x1,y1)
            x2=x+math.cos(endparam)
            y2=y+math.sin(endparam)
            endpoint=factory2d.CreatePoint(x2,y2)
            circle2d.StartPoint=startpoint
            circle2d.EndPoint=endpoint
            factory2d=mysketch.CloseEdition()
            rootpart.Update()
            return f"creat a 2D circle{circle2d.Name} ({x},{y},{r}), startpoint{startpoint.Name}:({x1},{y1}), endpoint{endpoint.Name}:({x2},{y2}) in {mysketch.Name}"
        except:
            factory2d=mysketch.CloseEdition()
            rootpart.Update()
            return "no startpoint and endpoint for the circle curve"
    except:
        factory2d=mysketch.CloseEdition()
        rootpart.Update()
        return "can't create a 2d circle"

@mcp.tool()#2026.4.26,2026.4.28,2026.5.24
async def create_line2d(body_name:str,sketch_name:str,x1:float,y1:float,x2:float,y2:float)->str:
    '''create a line in sketch

        Args:
        body_name: the name of the body
        sketch_name: the name of the sketch
        x1: The X coordinate of line first extremity
        y1: The Y coordinate of line first extremity
        x2: The X coordinate of line second extremity
        y2: The Y coordinate of line second extremity
    '''
    global catia
    rootpart=catia.ActiveDocument.Part
    mybody=rootpart.Bodies.Item(body_name)
    try:
        mysketch=mybody.Sketches.Item(sketch_name)
    except:
        return f"can't find the sketch({sketch_name}) in body({body_name})"
    try:    
        factory2d=mysketch.OpenEdition()
        line2d=factory2d.CreateLine(x1,y1,x2,y2)
        try:
            startpoint=factory2d.CreatePoint(x1,y1)
            endpoint=factory2d.CreatePoint(x2,y2)
            line2d.StartPoint=startpoint
            line2d.EndPoint=endpoint
            factory2d=mysketch.CloseEdition()
            rootpart.Update()
            return f"create a line {line2d.Name}, startpoint{startpoint.Name}:({x1},{y1}), endpoint{endpoint.Name}:({x2},{y2}) in {mysketch.Name}"
        except:
            factory2d=mysketch.CloseEdition()
            rootpart.Update()
            return "no startpoint and endpoint for the line"
    except:
        factory2d=mysketch.CloseEdition()
        rootpart.Update()
        return "can't create a line in sketch"
    
"""@mcp.tool()#2026.4.29,2026.5.24
def create_controlpoint2d(body_name:str,sketch_name:str,x:float,y:float,curvature:float,vx:float,vy:float)->str:
    '''create a control point

        Args:
        body_name: the name of the body
        sketch_name: the name of the sketch
        x: The X coordinate of point to create
        y: The Y coordinate of point to create
        curvature: the curvature properties of the spline control point
        vx: The X coordinate of the unit-vector for curvature direction at the point
        vy: The Y coordinate of the unit-vector for curvature direction at the point
    '''
    global catia
    rootpart=catia.ActiveDocument.Part
    try:
        mysketch=mybody.Sketches.Item(sketch_name)
    except:
        return f"can't find the sketch({sketch_name}) in body({body_name})"
    factory2d=mysketch.OpenEdition()
    mycontrolpoint=factory2d.CreateControlPoint(x,y)
    try:
        mycontrolpoint.SetTangent(vx,vy)
        mycontrolpoint.Curvature=curvature
        factory2d=mysketch.CloseEdition()
        rootpart.Update()
        return f"create a control point{mycontrolpoint.Name}:({x},{y},{curvature})"
    except:
        factory2d=mysketch.CloseEdition()
        rootpart.Update()
        return "can't add curvature in control point"
"""

@mcp.tool()#2026.4.26-2026.4.28，2026.4.29,2026.5.2,2026.5.24
async def add_monoeltcst(csttype:int,elem_name:str,sketch_name:str,body_name:str)->str:
    '''add a constraint in a sketch for one element

        Args:
        csttype: the value of CatConstraintType, that can be in tool(enum_CatConstraintType) found
        elem_name: the name of the geometric element
        sketch_name: the name of the sketch
        body_name: the name of the body   
    '''
    global catia
    rootpart=catia.ActiveDocument.part
    mybody=rootpart.Bodies.Item(body_name)
    try:
        mysketch=mybody.Sketches.Item(sketch_name)
    except:
        return f"can't find the sketch({sketch_name}) in body({body_name})"
    try:
        myelem=mysketch.GeometricElements.Item(elem_name)
    except:
        return f"can't find the element in {mysketch.Name}"

    try:    
        factory2d=mysketch.OpenEdition()
        mycontraint=mysketch.Constraints.AddMonoEltCst(csttype,myelem)
        factory2d=mysketch.CloseEdition()
        rootpart.Update()
        return f"added a constraint for {myelem.Name}"
    except:
        factory2d=mysketch.CloseEdition()
        rootpart.Update()
        return "can't add a constraint"  

@mcp.tool()#2026.4.28,2026.4.29,2026.5.2,2026.5.24,2026.5.30
async def add_bieltcst(csttype:int,firstelem_name:str,secondelem_name:str,sketch_name:str,body_name:str)->str:
    '''add a constraint in a sketch for two element

        Args:
        csttype: the value of CatConstraintType, that can be in tool(enum_CatConstraintType) found
        firstelem_name: the name of the geometric element
        secondelem_name: the name of the geometric element
        sketch_name: the name of the sketch
        body_name: the name of the body   
    '''
    global catia
    rootpart=catia.ActiveDocument.part
    mybody=rootpart.Bodies.Item(body_name)
    try:
        mysketch=mybody.Sketches.Item(sketch_name)
    except:
        return f"can't find the sketch({sketch_name}) in body({body_name})"
    try:
        if firstelem_name=="VDirection":
            myfirstelem=mysketch.GeometricElements.Item(1).VerticalReference
        elif firstelem_name=="HDirection":
            myfirstelem=mysketch.GeometricElements.Item(1).HorizontalReference
        else:
            myfirstelem=mysketch.GeometricElements.Item(firstelem_name)
        if secondelem_name=="VDirection":
            mysecondelem=mysketch.GeometricElements.Item(1).VerticalReference
        elif secondelem_name=="HDirection":
            mysecondelem=mysketch.GeometricElements.Item(1).HorizontalReference
        else:
            mysecondelem=mysketch.GeometricElements.Item(secondelem_name)
    except:
        return f"can't find the two elements in {mysketch.Name}"

    try:    
        factory2d=mysketch.OpenEdition()
        mycontraint=mysketch.Constraints.AddBiEltCst(csttype,myfirstelem,mysecondelem)
        factory2d=mysketch.CloseEdition()
        rootpart.Update()
        return f"added a constraint for {myfirstelem.Name} and {mysecondelem.Name}"
    except:
        factory2d=mysketch.CloseEdition()
        rootpart.Update()
        return "can't add a constraint"

@mcp.tool()#2026.4.30,2026.5.2,2026.5.24,2026.5.30
async def add_trieltcst(csttype:int,firstelem_name:str,secondelem_name:str,thirdelem_name:str,sketch_name:str,body_name:str)->str:
    '''add a constraint in a sketch for two element

        Args:
        csttype: the value of CatConstraintType, that can be in tool(enum_CatConstraintType) found
        firstelem_name: the name of the geometric element
        secondelem_name: the name of the geometric element
        thirdelem_name: the name of the geometric element
        sketch_name: the name of the sketch
        body_name: the name of the body   
    '''
    global catia
    rootpart=catia.ActiveDocument.part
    mybody=rootpart.Bodies.Item(body_name)
    try:
        mysketch=mybody.Sketches.Item(sketch_name)
    except:
        return f"can't find the sketch({sketch_name}) in body({body_name})"
    try:
        if firstelem_name=="VDirection":
            myfirstelem=mysketch.GeometricElements.Item(1).VerticalReference
        elif firstelem_name=="HDirection":
            myfirstelem=mysketch.GeometricElements.Item(1).HorizontalReference
        else:
            myfirstelem=mysketch.GeometricElements.Item(firstelem_name)
        if secondelem_name=="VDirection":
            mysecondelem=mysketch.GeometricElements.Item(1).VerticalReference
        elif secondelem_name=="HDirection":
            mysecondelem=mysketch.GeometricElements.Item(1).HorizontalReference
        else:
            mysecondelem=mysketch.GeometricElements.Item(secondelem_name)
        if thirdelem_name=="VDirection":
            mythirdelem=mysketch.GeometricElements.Item(1).VerticalReference
        elif thirdelem_name=="HDirection":
            mythirdelem=mysketch.GeometricElements.Item(1).HorizontalReference
        else:
            mythirdelem=mysketch.GeometricElements.Item(thirdelem_name)
    except:
        return f"can't find the three elements in {mysketch.Name}"

    try:    
        factory2d=mysketch.OpenEdition()
        mycontraint=mysketch.Constraints.AddTriEltCst(csttype,myfirstelem,mysecondelem,mythirdelem)
        factory2d=mysketch.CloseEdition()
        rootpart.Update()
        return f"added a constraint for {myfirstelem.Name},{mysecondelem.Name} and {mythirdelem.Name}"
    except:
        factory2d=mysketch.CloseEdition()
        rootpart.Update()
        return "can't add a constraint"

@mcp.tool()#2026.4.26,2026.4.29，2026.5.24
async def add_pad(body_name:str,sketch_name:str,h: float,pad_name:str)->str:
    '''add a new pad

        Args:
        body_name: the name of the body 
        sketch_name: the name of the sketch
        h: the long/height for the pad
        pad_name:the name for the new pad
    '''
    global catia
    rootpart=catia.ActiveDocument.Part
    mybody=rootpart.Bodies.Item(body_name)
    try:
        mysketch=mybody.Sketches.Item(sketch_name)
    except:
        return f"can't find the sketch({sketch_name}) in body({body_name})"
    #虽然pad是跟随着sketch同属于同一个Body，但这一步还是要有，保持逻辑清晰，从属明确
    try:
        newpad=rootpart.ShapeFactory.AddNewPad(mysketch,h)
    except:
        return "can't add a pad"
    else:
        rootpart.Update()
        newpad.Name=pad_name
        return f"create a pad:{newpad.Name}"

@mcp.tool()#2026.5.24
async def add_pocket(body_name:str,sketch_name:str,h:float,pocket_name:str)->str:
    '''add a new pocket

        Args:
        body_name: the name of the body
        sketch_name: the name of the sketch
        h: the depth for the pocket
        pad_name:the name for the new pad
    '''
    global catia
    rootpart=catia.ActiveDocument.Part
    mybody=rootpart.Bodies.Item(body_name)
    try:
        mysketch=mybody.Sketches.Item(sketch_name)
    except:
        return f"can't find the sketch({sketch_name}) in body({body_name})"
    #虽然pocket是跟随着sketch同属于同一个Body，但这一步还是要有，保持逻辑清晰，从属明确
    try:
        newpocket=rootpart.ShapeFactory.AddNewPocket(mysketch,h)
    except:
        return "can't add a pocket"
    else:
        rootpart.Update()
        newpocket.Name=pocket_name
        return f"create a pocket:{newpocket.Name}"

@mcp.tool()#2026.5.25,2026.5.30
async def add_shaft(body_name:str, sketch_name:str,shaft_name:str,rotationaxis_name:str)->str:
    '''add a new shaft

        Args:
        body_name: the name of the body
        sketch_name: the name of the sketch
        shaft_name:the name for the new shaft
        rotationaxis_name: the name of the rotation axis
    '''
    global catia
    rootpart=catia.ActiveDocument.Part
    mybody=rootpart.Bodies.Item(body_name)
    try:
        mysketch=mybody.Sketches.Item(sketch_name)
    except:
        return f"can't find the sketch({sketch_name}) in body({body_name})"
    #虽然shaft是跟随着sketch同属于同一个Body，但这一步还是要有，保持逻辑清晰，从属明确
    try:
        #用空引用建空壳shaft
        emptyref=rootpart.CreateReferenceFromName("")
        newshaft=rootpart.ShapeFactory.AddNewShaftFromRef(emptyref)
    except:
        return "can't create a empty Reference"
    else:
        try:
            #设轮廓
            profileref=rootpart.CreateReferenceFromObject(mysketch)
            newshaft.SetProfileElement(profileref)
        except:
            return "can't set the profile for shaft"
        else:
            try:
                #设旋转轴
                if rotationaxis_name=="VDirection":
                    axis2d=mysketch.GeometricElements.Item("AbsoluteAxis")
                    axisref=axis2d.GetItem("VDirection")
                elif rotationaxis_name=="HDirection":
                    axis2d=mysketch.GeometricElements.Item("AbsoluteAxis")
                    axisref=axis2d.GetItem("HDirection")
                else:
                    axisref=mysketch.GeometricElements.Item(rotationaxis_name)

                newshaft.RevoluteAxis=axisref
            except:
                return "can't set axis for shaft"
            else:    
                rootpart.Update()
                newshaft.Name=shaft_name
                return f"create a shaft:{shaft_name}"
    
@mcp.tool()#2026.5.30
async def add_groove(body_name:str,sketch_name:str,groove_name:str,rotationaxis_name:str)->str:
    '''add a new groove

        Args:
        body_name: the name of the body
        sketch_name: the name of the sketch
        groove_name:the name for the new groove
        rotationaxis_name: the name of the rotation axis
    '''
    global catia
    rootpart=catia.ActiveDocument.Part
    mybody=rootpart.Bodies.Item(body_name)
    try:
        mysketch=mybody.Sketches.Item(sketch_name)
    except:
        return f"can't find the sketch({sketch_name}) in body({body_name})"
    #虽然groove是跟随着sketch同属于同一个Body，但这一步还是要有，保持逻辑清晰，从属明确
    try:
        #用空引用建空壳groove
        emptyref=rootpart.CreateReferenceFromName("")
        newgroove=rootpart.ShapeFactory.AddNewGrooveFromRef(emptyref)
    except:
        return "can't create a empty Reference"
    else:
        try:
            #设轮廓
            profileref=rootpart.CreateReferenceFromObject(mysketch)
            newgroove.SetProfileElement(profileref)
        except:
            return "can't set the profile for groove"
        else:
            try:
                #设旋转轴
                if rotationaxis_name=="VDirection":
                    axis2d=mysketch.GeometricElements.Item("AbsoluteAxis")
                    axisref=axis2d.GetItem("VDirection")
                elif rotationaxis_name=="HDirection":
                    axis2d=mysketch.GeometricElements.Item("AbsoluteAxis")
                    axisref=axis2d.GetItem("HDirection")
                else:
                    axisref=mysketch.GeometricElements.Item(rotationaxis_name)
                
                newgroove.RevoluteAxis=axisref
            except:
                return "can't set axis for groove"
            else:    
                rootpart.Update()
                newgroove.Name=groove_name
                return f"create a groove:{groove_name}"

@mcp.tool()#2026.5.31-2026.6.3, 2026.6.20-2026.6.21
async def add_hole(
    body_name:str,hole_name:str,
    selectet:bool,x:float,y:float,z:float,
    hole_type:int,hole_anchormode:int,a2:float,d2:float,l2:float,countersunk_mode:int,
    hole_bottomtype:int,bottomlimit_mode:int,d1:float,a1:float,l1:float,
    hole_threadingmode:int,hole_threadside:int,d3:float,l3:float,p3:float
    )->str:
    '''add a new hole

        Args:
        body_name: the name of the body
        hole_name:the name for the new hole
        selectet: the user have to select a plane for create a hole; "True" means the user have already selected a plane, otherwise, it is "False".
        x: Origin point x absolute coordinate
        y: Origin point y absolute coordinate
        z: Origin point z absolute coordinate
        Sets the origin point which the hole is anchored to. If mandatory, the entry point will be projected onto a tangent plane.
        hole_type: the value of CatHoleType, that can be in tool(enum_CatHoleType) found.
        hole_anchormode: the value of CatHoleAnchorMode, that can be in tool(enum_CatHoleAnchorMode) found; it's pertinent when the hole type is: Counterbored or Counterdrilled
        a2: the value of the head angle; it's valid when the hole type is: Tapered or Counterdrilled or Countersunk
        d2: the value of the head diameter; it's valid when the hole type is: Counterbored or Counterdrilled
        l2: the value of the head depth; it's valid when the hole type is: Counterbored or Counterdrilled or Countersunk
        countersunk_mode: the value of CatCSHoleMode, that can be in tool(enum_CatCSHoleMode) found.
        hole_bottomtype: the value of CatHoleBottomType, that can be in tool(enum_CatHoleBottomType) found
        bottomlimit_mode: the value of CatLimitMode, that can be in tool(enum_CatLimitMode) found.
        d1: the value of the hole diameter
        a1: the value of the hole bottom angle
        l1: The value of the hole depth.
        hole_threadingmode: the value of CatHoleThreadingMode, that can in tool(enum_CatHoleThreadingMode) found.
        hole_threadside: the value of CatHoleThreadSide, that can in tool(enum_CatHoleThreadSide) found; it's valid when the hole is threaded.
        d3: the value of the thread diameter; it's valid when the hole is threaded, and the thread diameter must bigger than the hole diameter.
        l3: the value of the thread depthe; it's valid when the hole is threaded, and the thread depth is not bigger than the hole depth.
        p3: the value of the thread pitch; it's valid when the hole is threaded.
    '''
    if selectet==True:
        global catia
        mydoc=catia.ActiveDocument
        rootpart=mydoc.Part
        try:
            selection=mydoc.Selection
        except:
            return "don't get the Selection"
        if selection.Count==1:
            brepname=selection.Item(1).Value.Name
        else:
            return "please choose the plane again"
        
        s1=brepname.find("Brp:(")
        s2=brepname.index(")",s1)
        i=brepname[s1:s2+1]
        boundary_parentname=re.search(r"\((.+?);",i).group(1)
        boundary_index=re.search(r";(.+?)\)",i).group(1)
        brepname2=f"FSur:(Face:(Brp:({boundary_parentname};{boundary_index});None:();Cf14:());WithTemporaryBody;WithoutBuildError;WithInitialFeatureSupport;MonoFond;MFBRepVersion_CXR29)"
        #这里的MFBRepVersion_CXR29可能会出错，可用Macros来检验
        mybody=rootpart.Bodies.Item(body_name)
        try:
            myshape=mybody.Shapes.Item(boundary_parentname)
        except:
            return f"can't find the shape({boundary_parentname}) in body({body_name})"

        reference1=rootpart.CreateReferenceFromBRepName(brepname2,myshape)
        myhole=rootpart.ShapeFactory.AddNewHoleFromPoint(x,y,z,reference1,l1)
        
        #Extension延伸
        #BottomLimit.LimitMode目前只开发三种模式
        if bottomlimit_mode==0:
            myhole.BottomLimit.Dimension.Value=l1
            if hole_bottomtype==2:
                rootpart.HybridShapeFactory.DeleteObjectForDatum(myhole)
                rootpart.Update()
                return "The  trimmed hole bottom is not for blind bottomlimit"
            else:
                myhole.BottomType=hole_bottomtype
                if hole_bottomtype==1:
                    myhole.BottomAngle.Value=a1
        elif bottomlimit_mode==1 or bottomlimit_mode==2:
            if not hole_bottomtype==2:
                rootpart.HybridShapeFactory.DeleteObjectForDatum(myhole)
                rootpart.Update()
                return "the bottomlimit limitmode is not right for the bottomtype"
            else:
                myhole.BottomType=hole_bottomtype
        myhole.BottomLimit.LimitMode=bottomlimit_mode
        myhole.Diameter.Value=d1

        #Type类型
        if hole_type==0:
            myhole.Type=hole_type
        elif hole_type==1:
            myhole.Type=hole_type
            myhole.HeadAngle.Value=a2
            myhole.AnchorMode=hole_anchormode
        elif hole_type==2:
            myhole.Type=hole_type
            myhole.HeadDiameter.Value=d2
            myhole.HeadDepth.Value=l2
            myhole.AnchorMode=hole_anchormode
        elif hole_type==3:
            myhole.Type=hole_type
            if countersunk_mode==0:
                myhole.HeadDepth.Value=l2
                myhole.HeadAngle.Value=a2
            elif countersunk_mode==1:
                myhole.HeadDepth.Value=l2
                myhole.HeadDiameter.Value=d2
            elif countersunk_mode==2:
                myhole.HeadAngle.Value=a2
                myhole.HeadDiameter.Value=d2
        elif hole_type==4:
            myhole.Type=hole_type
            myhole.CounterDrilledMode=0
            #这里把CounterDrilledMode锁定为catCDModeNoCountersunkDiameter；是因为：Catia V5文档没有这个属性，且相关的值的变动Macros也展示不清晰
            myhole.HeadDiameter.Value=d2
            myhole.HeadDepth.Value=l2
            myhole.HeadAngle.Value=a2
            myhole.AnchorMode=hole_anchormode

        #Thread Definition螺纹定义
        if hole_threadingmode==1:
            myhole.ThreadingMode=hole_threadingmode
        elif hole_threadingmode==0:
            myhole.ThreadingMode=hole_threadingmode
            myhole.ThreadSide=hole_threadside
            if d3<=d1 or l3>l1:
                rootpart.HybridShapeFactory.DeleteObjectForDatum(myhole)
                rootpart.Update()
                return "the thread diameter or depth is not right"
            myhole.ThreadDiameter.Value=d3
            myhole.ThreadDepth.Value=l3
            myhole.ThreadPitch.Value=p3
        
        
        #我觉得无需开发
        '''#修改hole位置暂不开发
        mysketch=myhole.Sketch
        rootpart.InWorkObject=mysketch
        factory2d=mysketch.OpenEdition()
        '''
        rootpart.Update()
        myhole.Name=hole_name
        return "the hole is add"
    else:
        return "please in catia choose a plane for the new hole"

"""为检验add_hole的selection
@mcp.tool()#2026.6.2
async def check_selection()->str:
    '''check the selection
    '''
    global catia
    mydoc=catia.ActiveDocument
    rootpart=mydoc.Part
    try:
        selection=mydoc.Selection
    except:
        return "don't get the Selection"
    if selection.Count==1:
        brepname=selection.Item(1).Value.Name
        reference1=rootpart.CreateReferenceFromName(brepname)
        return f"brepname  {brepname}  parent  {reference1.Parent.Name}"
    else:
        return "please in catia choose a plane for the new hole"
"""

@mcp.tool()#2026.4.30, 2026.6.19
async def catiaquit(user_order:bool)->str:
    '''close catia

    '''
    global catia
    if user_order:
        catia.quit()
        return "Catia is closed"
    else:
        return "without user order"

@mcp.tool(title="ActiveDocument-Name")#2026.5.2
async def getname_activedoc()->str:
    '''get the name of active Document
    '''
    global catia
    mydoc=catia.ActiveDocument
    return f"the name for the active document is: {mydoc.Name}"

@mcp.tool(title="PartDocument-Properties")#2026.5.1,2026.5.25
async def info_partdoc(partdoc_name:str)->str:
    '''Properties for the part document

        Args:
        partdoc_name: the name of the part document, with .CATPart
    '''
    global catia
    mydoc=catia.Documents.Item(partdoc_name)
    return f"Document:{mydoc.Name}\nLocation:{mydoc.Path}\n{mydoc.FullName}\nonly for read:{mydoc.ReadOnly}\nhasn't changed:{mydoc.Saved}\n"

@mcp.tool(title="Part-Properties")#2026.5.2,2026.5.6,2026.5.25
async def info_rootpart(partdoc_name:str)->str:
    '''Properties for the rootpart

        Args:
        partdoc_name: the name of the part document, with .CATPart    
    '''
    global catia
    mydoc=catia.Documents.Item(partdoc_name)
    rootpart=mydoc.Part
    sentence=f"part: {rootpart.Name} form {mydoc.Name}\nThe origine Planes: PlaneXY, PlaneYZ, PlaneZX\nThe number of axissystems: {rootpart.AxisSystems.Count}\ntolerant body: Mainbody, the number of bodies:{rootpart.Bodies.Count}\nThe number of constraints: {rootpart.Constraints.Count}\n"
    if rootpart.Bodies.Count==0:
        return " don't have MainBody"
    else:
        sentence1=None
        for i in range(1,rootpart.Bodies.Count+1):
            mybody=rootpart.Bodies.Item(i)
            sentence1=f" {i}. Body name : {mybody.Name}\n"
            sentence=sentence+sentence1
    return sentence

@mcp.tool(title="Body-Properties")#2026.5.3-2026.5.6,2026.5.26
async def info_body(partdoc_name:str,body_name:str)->str:
    '''Properties for the body
    
        Args:
        partdoc_name: the name of the part, with .CATPart
        body_name: the name of the body
    '''
    global catia
    mydoc=catia.Documents.Item(partdoc_name)
    try:
        mybody=mydoc.Part.Bodies.Item(body_name)
        sentence=f"Body: {body_name} form {partdoc_name}\nThe number of shapes: {mybody.Shapes.Count}\n"
    except:
        return f"can't find body {body_name} in part {partdoc_name}"
    else:
        if mybody.Shapes.Count>0:
            sentence1=None
            for i in range(1,mybody.Shapes.Count+1):
                myshape=mybody.Shapes.Item(i)
                sentence1=f" {i}. Shape name: {myshape.Name}\n"
                sentence=sentence+sentence1
        sentence=sentence+f"The number of sketches: {mybody.Sketches.Count}"
        if mybody.Sketches.Count>0:
            sentence2=None
            for i in range(1,mybody.Sketches.Count+1):
                mysketch=mybody.Sketches.Item(i)
                sentence2=f" {i}. Sketch name: {mysketch.Name}\n"
                sentence=sentence+sentence2
        return sentence
    
@mcp.tool(title="Sketch-Properties")#2026.5.26
async def info_sketch(partdoc_name:str,body_name:str,sketch_name:str)->str:
    '''Properties for sketch

    Args:
        partdoc_name: the name of the part document,with .CATPart
        body_name: the name of the body
        sketch_name: the name of the sketch
    '''
    global catia
    mydoc=catia.Documents.Item(partdoc_name)
    rootpart=mydoc.Part
    try:
        mybody=rootpart.Bodies.Item(body_name)
    except:
        return f"can't find body {body_name} in part {partdoc_name}"
    else:
        try:
            mysketch=mybody.Sketches.Item(sketch_name)
        except:
            return f"can't find sketch {sketch_name} in body {body_name}"
        else:
            factory2d=mysketch.OpenEdition()         
            sentence=f"sketch: {sketch_name}\n"
            cst_number=mysketch.Constraints.Count
            sentence1=f"Constraints number: {cst_number}\n"
            if cst_number>0:
                for i in range (1,cst_number+1):
                    mycst=mysketch.Constraints.Item(i)
                    sentence1=sentence1+f"{i}. Constraint name: {mycst.Name}\n"        
            sentence=sentence+sentence1
            geoel_number=mysketch.GeometricElements.Count
            sentence2=f"Geometric elements number: {geoel_number}\n"
            if geoel_number>0:
                for i in range(1,geoel_number+1):
                    mygeoel=mysketch.GeometricElements.Item(i)
                    if i==1:
                        try:
                            sentence2=sentence2+f"{i}. Absolute axis is made of: {mygeoel.Origin.Name},{mygeoel.HorizontalReference.Name},{mygeoel.VerticalReference.Name}\n"
                        except:
                            return " Absolution axis fehler"
                    else:
                        sentence2=sentence2+f"{i}. Geometrical element name: {mygeoel.Name}\n"        
            sentence=sentence+sentence2
            factory2d=mysketch.CloseEdition()
            return sentence

@mcp.tool(title="Point2D-Properties")#2026.5.28-2026.5.29
async def info_point2d(partdoc_name:str,body_name:str,sketch_name:str,point_name:str)->str:
    '''Properties for point

        Args:
        partdoc_name: the name of the part document,with .CATPart
        body_name: the name of the body
        sketch_name: the name of the sketch
        point_name: the name of the point
    '''
    global catia
    mydoc=catia.Documents.Item(partdoc_name)
    rootpart=mydoc.Part
    try:
        mybody=rootpart.Bodies.Item(body_name)
    except:
        return f"can't find body {body_name} in part {partdoc_name}"
    else:
        try:
            mysketch=mybody.Sketches.Item(sketch_name)
        except:
            return f"can't find sketch {sketch_name} in body {body_name}"
        else:
            try:
                factory2d=mysketch.OpenEdition()
                mypoint=mysketch.GeometricElements.Item(point_name)
            except:
                return f"can't find point {point_name} in sketch {sketch_name}"
            else:
                '''
                #约束只显示绝对值，没有负值。虽不能用却是好的经验
                vdirection=mysketch.GeometricElements.Item(1).VerticalReference
                mycontraint=mysketch.Constraints.AddBiEltCst(1,mypoint,vdirection)
                x=mycontraint.Dimension.ValueAsString()
                hdirection=mysketch.GeometricElements.Item(1).HorizontalReference
                mycontraint1=mysketch.Constraints.AddBiEltCst(1,mypoint,hdirection)
                y=mycontraint1.Dimension.ValueAsString()
                factory2d=mysketch.CloseEdition()
                rootpart.Update()
                #由于point2d没有暴露属性的接口，GetCoordinates()是无效输出，他把列表输出为元组，但是数据没改变，所以是无效的，应该是这个函数的逻辑问题；所以不得已使用约束来查看，上段代码可根据其接口的修正而改变
                factory2d=mysketch.OpenEdition()
                selection=mydoc.Selection
                selection.Clear()
                selection.Add(mycontraint)
                selection.Delete()
                selection.Clear()
                selection.Add(mycontraint1)
                selection.Delete()
                factory2d=mysketch.CloseEdition()
                rootpart.Update()
                #由于Constraints.Remove()是无效的，所以只能用selection()，上段代码可根据前者的修正而改变
                return f"the point is ({x},{y})"
                '''
                vba_code = """
                    Public Function get_coordinates(point2_d)
                    Dim oPoint(1)
                    point2_d.GetCoordinates oPoint
                    get_coordinates = oPoint
                    End Function
                    """

                system_service = catia.SystemService
                result = system_service.Evaluate(vba_code, 0, "get_coordinates", [mypoint])
                x=result[0]
                y=result[1]
                return f"point {point_name}: ({x},{y})"

@mcp.tool(title="Line2D-Properties")#2026.5.29
async def info_line2d(partdoc_name:str,body_name:str,sketch_name:str,line_name:str)->str:
    '''Properties for point

        Args:
        partdoc_name: the name of the part document,with .CATPart
        body_name: the name of the body
        sketch_name: the name of the sketch
        line_name: the name of the line
    '''
    global catia
    mydoc=catia.Documents.Item(partdoc_name)
    rootpart=mydoc.Part
    try:
        mybody=rootpart.Bodies.Item(body_name)
    except:
        return f"can't find body {body_name} in part {partdoc_name}"
    else:
        try:
            mysketch=mybody.Sketches.Item(sketch_name)
        except:
            return f"can't find sketch {sketch_name} in body {body_name}"
        else:
            try:
                factory2d=mysketch.OpenEdition()
                myline=mysketch.GeometricElements.Item(line_name)
            except:
                return f"can't find point {line_name} in sketch {sketch_name}"
            else:   
                try:
                    startpoint=myline.StartPoint
                except:
                    errorinfo="don't have startpoint"
                    try:
                        endpoint=myline.EndPoint
                        return "have endpoint"+errorinfo
                    except:
                        return "don't have endpoint"+errorinfo
                    finally:
                        factory2d=mysketch.CloseEdition()
                else:
                    errorinfo="have startpoint"
                try:
                    endpoint=myline.EndPoint
                except:
                    return "don't have endpoint"+errorinfo
                else:
                    return f"line {line_name}: startpoin {startpoint.Name} endpoint {endpoint.Name}"
                finally:
                    factory2d=mysketch.CloseEdition()
            
@mcp.tool(title="Circle-Properties")#2026.5.29                    
async def info_circle2d(partdoc_name:str,body_name:str,sketch_name:str,circle_name:str)->str:
    '''Properties for point

        Args:
        partdoc_name: the name of the part document,with .CATPart
        body_name: the name of the body
        sketch_name: the name of the sketch
        circle_name: the name of the circle
    '''
    global catia
    mydoc=catia.Documents.Item(partdoc_name)
    rootpart=mydoc.Part
    try:
        mybody=rootpart.Bodies.Item(body_name)
    except:
        return f"can't find body {body_name} in part {partdoc_name}"
    else:
        try:
            mysketch=mybody.Sketches.Item(sketch_name)
        except:
            return f"can't find sketch {sketch_name} in body {body_name}"
        else:
            try:
                factory2d=mysketch.OpenEdition()
                mycircle=mysketch.GeometricElements.Item(circle_name)
            except:
                return f"can't find circle {circle_name} in sketch {sketch_name}"
            else:
                try:
                    centerpoint=mycircle.CenterPoint
                except:
                    return "don't have centerpoint"
                else:
                    r=mycircle.Radius   
                    try:
                        startpoint=mycircle.StartPoint
                    except:
                        errorinfo="don't have startpoint"
                        try:
                            endpoint=mycircle.EndPoint
                            return "have endpoint"+errorinfo
                        except:
                            return f"{circle_name} don't have endpoint"+errorinfo +f", or is a close circle; center point {centerpoint.Name}, radius {r}"
                        finally:
                            factory2d=mysketch.CloseEdition()
                    else:
                        errorinfo="have startpoint"
                    try:
                        endpoint=mycircle.EndPoint
                    except:
                        return "don't have endpoint"+errorinfo
                    else:
                        return f"{circle_name}: endpoint {endpoint.Name}, startpoint {startpoint.Name}, center point {centerpoint.Name}, radius {r}"
                    finally:
                        factory2d=mysketch.CloseEdition()

@mcp.tool(title="Pad-Properties")#2026.6.4
async def info_pad(partdoc_name:str,body_name:str,pad_name:str)->str:
    '''Properties for pad

        Args:
        partdoc_name: the name of the part document,with .CATPart
        body_name: the name of the body
        pad_name: the name of the pad
    '''
    global catia
    mydoc=catia.Documents.Item(partdoc_name)
    rootpart=mydoc.Part
    try:
        mybody=rootpart.Bodies.Item(body_name)
    except:
        return f"can't find body {body_name} in part {partdoc_name}"
    else:
        try:
            mypad=mybody.Item(pad_name)
        except:
            return f"can't find pad {pad_name} in body {body_name}"
        else:
            mysketch=mypad.Sketch
            mylength=mypad.FirstLimit.Value
            return f"pad sketch: {mysketch.Name}; Length: {mylength}"

@mcp.tool(title="Pocket-Properties")#2026.6.4
async def info_pocket(partdoc_name:str,body_name:str,pocket_name:str)->str:
    '''Properties for pocket

        Args:
        partdoc_name: the name of the part document,with .CATPart
        body_name: the name of the body
        pocket_name: the name of the pocket
    '''
    global catia
    mydoc=catia.Documents.Item(partdoc_name)
    rootpart=mydoc.Part
    try:
        mybody=rootpart.Bodies.Item(body_name)
    except:
        return f"can't find body {body_name} in part {partdoc_name}"
    else:
        try:
            mypocket=mybody.Item(pocket_name)
        except:
            return f"can't find pad {pocket_name} in body {body_name}"
        else:
            mysketch=mypocket.Sketch
            mylength=mypocket.FirstLimit.Value
            return f"pocket sketch: {mysketch.Name}; Length: {mylength}"

@mcp.tool(title="Shaft-Properties")#2026.6.4
async def info_shaft(partdoc_name:str,body_name:str,shaft_name:str)->str:
    '''Properties for shaft

        Args:
        partdoc_name: the name of the part document,with .CATPart
        body_name: the name of the body
        shaft_name: the name of the shaft
    '''
    global catia
    mydoc=catia.Documents.Item(partdoc_name)
    rootpart=mydoc.Part
    try:
        mybody=rootpart.Bodies.Item(body_name)
    except:
        return f"can't find body {body_name} in part {partdoc_name}"
    else:
        try:
            myshaft=mybody.Item(shaft_name)
        except:
            return f"can't find pad {shaft_name} in body {body_name}"
        else:
            mysketch=myshaft.Sketch
            myfirstangle=myshaft.FirstAngle.Value
            mysecondangle=myshaft.SecondAngle.Value
            myrevaxis=myshaft.RevoluteAxis
            return f"shaft sketch: {mysketch.Name}\nFirst angle: {myfirstangle}\nSecond angle: {mysecondangle}\nRevolute axis: {myrevaxis.Name}"
        
@mcp.tool(title="Groove-Properties")#2026.6.4
async def info_groove(partdoc_name:str,body_name:str,groove_name:str)->str:
    '''Properties for groove

        Args:
        partdoc_name: the name of the part document,with .CATPart
        body_name: the name of the body
        groove_name: the name of the groove
    '''
    global catia
    mydoc=catia.Documents.Item(partdoc_name)
    rootpart=mydoc.Part
    try:
        mybody=rootpart.Bodies.Item(body_name)
    except:
        return f"can't find body {body_name} in part {partdoc_name}"
    else:
        try:
            mygroove=mybody.Item(groove_name)
        except:
            return f"can't find pad {groove_name} in body {body_name}"
        else:
            mysketch=mygroove.Sketch
            myfirstangle=mygroove.FirstAngle.Value
            mysecondangle=mygroove.SecondAngle.Value
            myrevaxis=mygroove.RevoluteAxis
            return f"groove sketch: {mysketch.Name}\nFirst angle: {myfirstangle}\nSecond angle: {mysecondangle}\nRevolute axis: {myrevaxis.Name}"

@mcp.tool(title="enum_CatConstraintType")#2026.4.29,2026.5.2
async def check_CatConstraintType()->dict:
    '''check the value of the CatConstraintType
    '''
    CatConstraintType={
    "catCstTypeReference":0,
    "catCstTypeDistance":1,
    "catCstTypeOn":2,
    "catCstTypeConcentricity":3,
    "catCstTypeTangency":4,
    "catCstTypeLength":5,
    "catCstTypeAngle":6,
    "catCstTypePlanarAngle":7,
    "catCstTypeParallelism":8,
    "catCstTypeAxisParallelism":9,
    "catCstTypeHorizontality":10,
    "catCstTypePerpendicularity":11,
    "catCstTypeAxisPerpendicularity":12,
    "catCstTypeVerticality":13,
    "catCstTypeRadius":14,
    "catCstTypeSymmetry":15,
    "catCstTypeMidPoint":16,
    "catCstTypeEquidistance":17,
    "catCstTypeMajorRadius":18,
    "catCstTypeMinorRadius":19,
    "catCstTypeSurfContact":20,
    "catCstTypeLinContact":21,
    "catCstTypePoncContact":22,
    "catCstTypeChamfer":23,
    "catCstTypeChamferPerpend":24,
    "catCstTypeAnnulContact":25,
    "catCstTypeCylinderRadius":26,
    "catCstTypeStContinuity":27,
    "catCstTypeStDistance":28,
    "catCstTypeSdContinuity":29,
    "catCstTypeSdShape":30}
    return CatConstraintType

@mcp.tool(title="enum_CatHoleType")#2026.6.1
async def check_CatHoleType()->dict:
    '''check the value of the CatHoleType
    '''
    CatHoleType={
        "catSimpleHole":0,
        "catTaperedHole":1,
        "catCounterboredHole":2,
        "catCountersunkHole":3,
        "catCounterdrilledHole":4
        }
    return CatHoleType

@mcp.tool(title="enum_CatHoleAnchorMode")#2026.6.1
async def check_CatHoleAnchorMode()->dict:
    '''check the value of the CatHoleAnchorMode
    '''
    CatHoleAnchorMode={
        "catExtremPointHoleAnchor":0,
        "catCATMiddlePointHoleAnchor":1
    }
    return CatHoleAnchorMode

@mcp.tool(title="enum_CatHoleBottomType")#2026.6.1
async def check_CatHoleBottomType()->dict:
    '''check the value of the CatHoleBottomType
    '''
    CatHoleBottomType={
        "catFlatHoleBottom":0,
        "catVHoleBottom":1,
        "catTrimmedHoleBottom":2
    }
    return CatHoleBottomType


@mcp.tool(title="enum_CatLimitMode")#2026.6.1, 2026.6.20
async def check_CatLimitMode()->dict:
    '''check the value of the CatLimitMode
    '''
    #目前只开发三种
    CatLimitMode={
        "catOffsetLimit":0,
        "catUpToNextLimit":1,
        "catUpToLastLimit":2,
        #"catUptoPlaneLimit":3,
        #"catUpToSurfaceLimit":4,
        #"catUpThruNextLimit":5
    }
    
    return CatLimitMode

@mcp.tool(title="enum_CatHoleThreadingMode")#2026.6.1
async def check_CatHoleThreadingMode()->dict:
    '''check the value of the CatHoleThreadingMode
    '''
    CatHoleThreadingMode={
        "catThreadedHoleThreading":0,
        "catSmoothHoleThreading":1
    }
    return CatHoleThreadingMode

@mcp.tool(title="enum_CatHoleThreadSide")#2026.6.2
async def check_CatHoleThreadSide()->dict:
    '''check the value of the CatHoleThreadSide
    '''
    CatHoleThreadSide={
        "catRightThreadSide":0,
        "catLeftThreadSide":1
    }
    return CatHoleThreadSide

@mcp.tool(title="enum_CatCSHoleMode")#2026.6.21
async def check_CatCSHoleMode()->dict:
    '''check the value of the CatCSHoleMode
    '''
    CatCSHoleMode={
        "catCSModeDepthAngle":0,
        "catCSModeDepthDiameter":1,
        "catCSModeAngleDiameter":2
    }
    return CatCSHoleMode

@mcp.tool(title="Captureviewer")#2026.6.20
async def get_capture():
    '''check the view of the catia viewer
    '''
    global catia
    global CATIAMCP_FILE
    if not CATIAMCP_FILE:
        return "don't have file to save picture"
    else:
        viewer=catia.ActiveWindow.ActiveViewer
        viewer.CaptureToFile(5,f"{CATIAMCP_FILE}\\MyImage.jpeg")

    return Image(path=f"{CATIAMCP_FILE}\\MyImage.jpeg")



#mcpresource-unavailable

#mcpprompt-unavailable



#Initialize and run the server 2026.4.21
def main():
    mcp.run(transport="stdio")

if __name__=="__main__":
    main()
