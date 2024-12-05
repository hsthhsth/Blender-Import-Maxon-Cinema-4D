bl_info = {
    "name": "导入.c4d文件",
    "description": "导入Maxon Cinema 4D(.c4d)文件",
    "author": "红糖hsth",
    "version": (1, 0, 0),
    "category": "Import-Export",
    "warning": "须选择Maxon Cinema 4D安装文件夹路径后使用！",
    "doc_url": "https://pd.qq.com/s/51buqixcj",
    "tracker_url": "https://qm.qq.com/q/bOHrwm7Lck&group_code=891468989",
}

import bpy
import os
import platform
import subprocess
import tempfile
from bpy.props import StringProperty, BoolProperty
from bpy.types import Operator, AddonPreferences, FileHandler
from bpy_extras.io_utils import ImportHelper

# 偏好设置类
class ImportC4DPreferences(AddonPreferences):
    bl_idname = __name__
    c4d_install_path: StringProperty(
        name="Cinema 4D安装路径",
        description="Maxon Cinema 4D安装文件夹路径",
        subtype='DIR_PATH',
        default=""
    )

    def draw(self, context):
        layout = self.layout
        layout.prop(self, "c4d_install_path")

# 导入操作类
class ImportC4D(Operator, ImportHelper):
    bl_idname = "import_scene.c4d"
    bl_label = "导入C4D（.c4d）"
    bl_description = "导入Maxon Cinema 4D(.c4d)文件"
    bl_options = {'PRESET', 'UNDO'}

    filename_ext = ".c4d"
    filter_glob: StringProperty(default="*.c4d", options={'HIDDEN'})
    filepath: StringProperty(subtype='FILE_PATH', options={'SKIP_SAVE'})

    import_models: BoolProperty(name="导入模型", default=True)
    import_materials: BoolProperty(name="导入材质", default=False)
    import_lights: BoolProperty(name="导入灯光", default=False)
    import_cameras: BoolProperty(name="导入相机", default=False)
    import_splines: BoolProperty(name="导入样条曲线", default=False)
    import_animations: BoolProperty(name="导入动画", default=False)
    
    def draw(self, context):
        layout = self.layout
        for prop_name in ['import_models', 'import_materials', 'import_lights', 'import_cameras', 'import_splines', 'import_animations']:
            layout.prop(self, prop_name)

    @classmethod
    def poll(cls, context):
        return True

    def execute(self, context):
        c4d_file_path = self.filepath

        if not c4d_file_path or not c4d_file_path.lower().endswith(".c4d"):
            self.report({'ERROR'}, "无效的文件路径或扩展名。")
            return {'CANCELLED'}

        temp_dir = tempfile.gettempdir()
        fbx_file_path = os.path.join(temp_dir, "exported_file.fbx")

        addon_prefs = context.preferences.addons[__name__].preferences
        c4d_install_path = addon_prefs.c4d_install_path.strip()

        # 定义一个字典，将操作系统名称映射到可执行文件名
        os_to_executable = {
            "Windows": "c4dpy.exe",
            "Darwin": "c4dpy",  # macOS
        }
        # 获取当前操作系统名称
        os_name = platform.system()
        # 检查当前操作系统是否在支持的列表中
        c4dpy_executable = os_to_executable.get(os_name)
        if not c4dpy_executable:
            self.report({'ERROR'}, "不支持的操作系统。")
            return {'CANCELLED'}

        c4dpy_path = os.path.join(c4d_install_path, c4dpy_executable)
        if not os.path.isfile(c4dpy_path):
            self.report({'ERROR'}, f"无法在{c4dpy_path}找到c4dpy。")
            return {'CANCELLED'}

        self.export_c4d_to_fbx(c4dpy_path, c4d_file_path, fbx_file_path)

        if not os.path.isfile(fbx_file_path):
            self.report({'ERROR'}, "FBX导出失败。")
            return {'CANCELLED'}

        bpy.ops.import_scene.fbx(filepath=fbx_file_path, use_image_search=True, use_custom_normals=True)

        self.cleanup_unwanted_objects()

        self.report({'INFO'}, "导出和导入成功！")
        return {'FINISHED'}

    def cleanup_unwanted_objects(self):
        import_options = {
            'import_models': ('MESH', self.delete_objects_of_type),
            'import_materials': (None, self.delete_materials),
            'import_lights': ('LIGHT', self.delete_objects_of_type),
            'import_cameras': ('CAMERA', self.delete_objects_of_type),
            'import_splines': ('CURVE', self.delete_objects_of_type),
            'import_animations': (None, self.delete_animations),
        }
        for option, (obj_type, delete_method) in import_options.items():
            if not getattr(self, option):
                if obj_type:
                    delete_method(obj_type)
                else:
                    delete_method()

    def delete_objects_of_type(self, obj_type):
        bpy.ops.object.select_all(action='DESELECT')
        for obj in bpy.context.scene.objects:
            if obj.type == obj_type:
                obj.select_set(True)
        bpy.ops.object.delete(use_global=False)

    def delete_animations(self):
        for obj in bpy.context.scene.objects:
            if obj.animation_data:
                obj.animation_data_clear()

    def delete_materials(self):
        for obj in bpy.context.scene.objects:
            if obj.data and hasattr(obj.data, 'materials'):
                obj.data.materials.clear()
        for material in bpy.data.materials:
            bpy.data.materials.remove(material)

    def export_c4d_to_fbx(self, c4dpy_path, c4d_file, fbx_file):
        script_content = f"""
import c4d

def export_to_fbx(file_in, file_out):
    doc = c4d.documents.LoadDocument(file_in, c4d.SCENEFILTER_OBJECTS | c4d.SCENEFILTER_MATERIALS)
    if not doc:
        raise RuntimeError("加载文档失败")
    c4d.documents.SetActiveDocument(doc)
    c4d.documents.SaveDocument(doc, file_out, c4d.SAVEDOCUMENTFLAGS_DONTADDTORECENTLIST, c4d.FORMAT_FBX_EXPORT)

export_to_fbx(r"{c4d_file}", r"{fbx_file}")
"""
        with tempfile.NamedTemporaryFile(delete=False, suffix=".py") as temp_script_file:
            temp_script_file.write(script_content.encode('utf-8'))
            temp_script_file_path = temp_script_file.name
        try:
            subprocess.run([c4dpy_path, temp_script_file_path], check=True, capture_output=True, text=True)
        except subprocess.CalledProcessError as e:
            self.report({'ERROR'}, "使用c4dpy导出FBX失败")
        finally:
            os.remove(temp_script_file_path)

def invoke(self, context, event):
    return self.invoke_popup(context)

#允许通过拖放操作在3D视图区域中导入C4D（.c4d）文件
class IO_FH_C4D(FileHandler):
    bl_idname = "IO_FH_C4D"
    bl_label = "Maxon Cinema 4D"
    bl_import_operator = "import_scene.c4d"
    bl_file_extensions = ".c4d"

    @classmethod
    def poll_drop(cls, context):
        return context.area.type == 'VIEW_3D'

#在Blender的“文件”>“导入”菜单中添加Maxon Cinema 4D（.c4d）选项
def menu_func_import(self, context):
    self.layout.operator(ImportC4D.bl_idname, text="Maxon Cinema 4D(.c4d)")

def register():
    bpy.utils.register_class(ImportC4DPreferences)
    bpy.utils.register_class(ImportC4D)
    bpy.utils.register_class(IO_FH_C4D)
    bpy.types.TOPBAR_MT_file_import.append(menu_func_import)

def unregister():
    bpy.utils.unregister_class(ImportC4DPreferences)
    bpy.utils.unregister_class(ImportC4D)
    bpy.utils.unregister_class(IO_FH_C4D)
    bpy.types.TOPBAR_MT_file_import.remove(menu_func_import)

if __name__ == "__main__":
    register()