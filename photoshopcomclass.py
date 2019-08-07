# libs\photoshopcomclass.py

from win32com.client import Dispatch
import logging

log = logging.getLogger(__name__)


class PhotoshopCOM(object):

    psDisplayNoDialogs = 3  # from enum PsDialogModes

    psSaveChanges = 1  # from enum PsSaveOptions
    psDoNotSaveChanges = 2  # from enum PsSaveOptions

    psSmartObjectLayer = 17  # from enum PsLayerKind
    psTextLayer = 2  # from enum PsLayerKind
    psNormalLayer = 1  # from enum PsLayerKind

    psTransparentPixels = 0  # from enum PsTrimType

    # preview_size = 800 # Width in pixels of preview image
    preview_quality = 12  # JPEG quality from 1 to 12

    def __init__(self, path):

        self.app = Dispatch('Photoshop.Application')
        self.doc = self.app.Open(path)
        # [layer object, layer name, layer kind(-1 if it is LayerSet), layer visible?]
        self.all_layers = []
        self.list_all_layers()

    def cTID(self, text):

        return self.app.CharIDToTypeID(text)

    def sTID(self, text):

        return self.app.StringIDToTypeID(text)

    def cTString(self, text):

        return self.app.TypeIDToStringID(self.cTID(text))

    def select_layer(self):

        # Select layer as Ctrl+LKM on layer
        desc1 = Dispatch('Photoshop.ActionDescriptor')
        ref1 = Dispatch('Photoshop.ActionReference')
        ref1.PutProperty(self.cTID('Chnl'), self.cTID('fsel'))
        desc1.PutReference(self.cTID('null'), ref1)
        ref2 = Dispatch('Photoshop.ActionReference')
        ref2.PutEnumerated(self.cTID('Chnl'), self.cTID('Chnl'), self.cTID('Trsp'))
        desc1.PutReference(self.cTID('T   '), ref2)
        self.app.ExecuteAction(cTID('setd'), desc1, self.psDisplayNoDialogs)

    def list_all_layers(self):

        for layer in self.doc.Layers:
            if layer.Typename == 'ArtLayer':
                self.all_layers.append([layer, layer.Name, layer.Kind, layer.Visible])
            else:
                self.all_layers.append([layer, layer.Name, -1, layer.Visible])
                for layer2 in self.doc.LayerSets[layer.Name].Layers:
                    if layer2.Typename == 'ArtLayer':
                        self.all_layers.append([layer2, layer2.Name, layer2.Kind, layer2.Visible])
                    else:
                        self.all_layers.append([layer2, layer2.Name, -1, layer2.Visible])
                        for layer3 in self.doc.LayerSets[layer.Name].LayerSets[layer2.Name].Layers:
                            if layer3.Typename == 'ArtLayer':
                                self.all_layers.append([layer3, layer3.Name, layer3.Kind, layer3.Visible])
                            else:
                                self.all_layers.append([layer3, layer3.Name, -1, layer3.Visible])
                                for layer4 in self.doc.LayerSets[layer.Name].LayerSets[layer2.Name].LayerSets[layer3.Name].Layers:
                                    if layer4.Typename == 'ArtLayer':
                                        self.all_layers.append([layer4, layer4.Name, layer4.Kind, layer4.Visible])
                                    else:
                                        self.all_layers.append([layer4, layer4.Name, -1, layer4.Visible])

    def set_active_layer(self, layer):

        self.doc.ActiveLayer = layer

    def change_text(self, layer, text):

        if layer.Typename == 'ArtLayer' and layer.Kind == self.psTextLayer:
            layer.TextItem.Contents = text.replace('/r', '\r')

    def change_text_size(self, layer, size):

        if layer.Typename == 'ArtLayer' and layer.Kind == self.psTextLayer:
            layer.TextItem.size = size

    def make_selection(self, layer):

        self.set_active_layer(layer)
        # Select layer as Ctrl+LKM on layer
        desc2 = Dispatch('Photoshop.ActionDescriptor')
        ref3 = Dispatch('Photoshop.ActionReference')
        ref3.PutProperty(self.cTID('Chnl'), self.cTID('fsel'))
        desc2.PutReference(self.cTID('null'), ref3)
        ref4 = Dispatch('Photoshop.ActionReference')
        ref4.PutEnumerated(self.cTID('Chnl'), self.cTID('Chnl'), self.cTID('Trsp'))
        desc2.PutReference(self.cTID('T   '), ref4)
        self.app.ExecuteAction(self.cTID('setd'), desc2, self.psDisplayNoDialogs)

    def change_smart_image(self, layer, imagepath):

        if layer.Typename == 'ArtLayer' and layer.Kind == self.psSmartObjectLayer:
            self.set_active_layer(layer)
            # Open it to edit(Smart Layers only)
            desc3 = Dispatch('Photoshop.ActionDescriptor')
            self.app.ExecuteAction(self.sTID('placedLayerEditContents'), desc3, self.psDisplayNoDialogs)
            # Replace layer image with image from file, will resize image to layer size!
            desc3 = Dispatch('Photoshop.ActionDescriptor')
            desc3.PutPath(self.sTID('null'), imagepath)
            desc3.PutEnumerated(self.sTID('freeTransformCenterState'), self.sTID('quadCenterState'), self.sTID('QCSAverage'))
            desc4 = Dispatch('Photoshop.ActionDescriptor')
            desc4.PutUnitDouble(self.sTID('horizontal'), self.sTID('pixelsUnit'), 0.000000)
            desc4.PutUnitDouble(self.sTID('vertical'), self.sTID('pixelsUnit'), 0.000000)
            desc3.PutObject(self.sTID('offset'), self.sTID('offset'), desc4)
            self.app.ExecuteAction(self.sTID('placeEvent'), desc3, self.psDisplayNoDialogs)
            # Rasterize inserted layer
            desc5 = Dispatch('Photoshop.ActionDescriptor')
            ref5 = Dispatch('Photoshop.ActionReference')
            ref5.PutEnumerated(self.cTID('Lyr '), self.cTID('Ordn'), self.cTID('Trgt'))
            desc5.PutReference(self.cTID('null'), ref5)
            self.app.ExecuteAction(self.sTID('rasterizeLayer'), desc5, self.psDisplayNoDialogs)
            # Hide template layers, that were there before
            for layer in self.app.ActiveDocument.ArtLayers:
                if layer.Name != self.app.ActiveDocument.ActiveLayer.Name:
                    layer.Visible = False
            # Save and close smart object
            self.close_and_save()

    def save_preview(self, path):

        # Resize to lower resolution and save JPG copy with the lowest quality for preview purposes
        # self.app.ActiveDocument.ResizeImage(self.preview_size)
        options = Dispatch('Photoshop.JPEGSaveOptions')
        options.EmbedColorProfile = False
        options.Quality = self.preview_quality
        # Save preview for FRONT and BACK main groups of layers if they exists
        try:
            self.doc.ActiveLayer = self.doc.Layers['FRONT']
            self.make_visible(self.doc.ActiveLayer)
            try:
                self.doc.ActiveLayer = self.doc.Layers['BACK']
                self.make_invisible(self.doc.ActiveLayer)
                self.app.ActiveDocument.SaveAs(path + 'front_preview.jpg', options, True)
                self.doc.ActiveLayer = self.doc.Layers['BACK']
                self.make_visible(self.doc.ActiveLayer)
                self.doc.ActiveLayer = self.doc.Layers['FRONT']
                self.make_invisible(self.doc.ActiveLayer)
                self.app.ActiveDocument.SaveAs(path + 'back_preview.jpg', options, True)
            except:
                self.app.ActiveDocument.SaveAs(path + 'front_preview.jpg', options, True)
        except:
            self.app.ActiveDocument.SaveAs(path + 'front_preview.jpg', options, True)

    def save_as(self, path):

        # Save As edited copy
        self.app.ActiveDocument.SaveAs(path)

    def close_no_save(self):

        # Close active doc WITHOUT saving
        self.app.ActiveDocument.Close(self.psDoNotSaveChanges)

    def close_and_save(self):

        # Close active doc WITH saving
        self.app.ActiveDocument.Close(self.psSaveChanges)

    def make_visible(self, layer):

        layer.Visible = True

    def make_invisible(self, layer):

        layer.Visible = False

    def trim_transparent(self):

        self.app.ActiveDocument.Trim(self.psTransparentPixels)

    def quit(self):
        self.app.ExecuteAction(self.cTID('quit'))
