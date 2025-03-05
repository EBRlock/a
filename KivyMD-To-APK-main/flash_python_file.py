from jnius import autoclass, cast

class Flash:
    @staticmethod
    def initialize():
        CameraManager = autoclass('android.hardware.camera2.CameraManager')
        Context=autoclass("android.content.Context")
        context=Context()
        self.camera_manager=cast("android.hardware.camera2.CameraManager",context.getSystemService(Context.CAMERA_SERVICE))
        cameraid = self.camera_manager.getCameraIdList()[0]
    @staticmethod
    def on():
        self.camera_manager.setTorchMode(Flash.cameraid, True)
    @staticmethod
    def off():
        self.camera_manager.setTorchMode(Flash.cameraid, False)
