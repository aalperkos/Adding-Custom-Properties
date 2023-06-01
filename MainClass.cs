using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Navisworks.Api;
using Autodesk.Navisworks.Api.Plugins;
using Autodesk.Navisworks.Api.Interop.ComApi;
using ComApiBridge = Autodesk.Navisworks.Api.ComApi.ComApiBridge;

namespace CustomPa
{
    [PluginAttribute("CustomPa_Add", "RollaBim", DisplayName = "CustomPa_Add", ToolTip = "Adding Custom Properties")]
    public class MainClass : AddInPlugin
    {
        public override int Execute(params string[] parameters)
        {
            // Current document (.NET)
            Document doc = Application.ActiveDocument;
            // Current document (COM)
            InwOpState10 cdoc = ComApiBridge.State;
            // Current selected items
            ModelItemCollection items = doc.CurrentSelection.SelectedItems;

            if (items.Count > 0)
            {
                // Input dialog
                InputDialog dialog = new InputDialog();
                dialog.ShowDialog();

                // Create a new Category (PropertyDataCollection)
                InwOaPropertyVec newcate = (InwOaPropertyVec)cdoc.ObjectFactory(nwEObjectType.eObjectType_nwOaPropertyVec, null, null);

                foreach (ModelItem item in items)
                {
                    // Convert ModelItem to COM Path
                    InwOaPath citem = (InwOaPath)ComApiBridge.ToInwOaPath(item);
                    // Get item's PropertyCategoryCollection
                    InwGUIPropertyNode2 cpropcates = (InwGUIPropertyNode2)cdoc.GetGUIPropertyNode(citem, true);
                    // Get PropertyCategoryCollection data
                    InwGUIAttributesColl propCol = cpropcates.GUIAttributes();
                    // Check if category name matches
                    bool categoryMatch = false;

                    foreach (InwGUIAttribute2 i in propCol)
                    {
                        if (i.UserDefined && i.ClassUserName == dialog.CategoryName)
                        {
                            // Category name matches
                            categoryMatch = true;

                            // a new propertycategory object
                            InwOaPropertyVec category = (InwOaPropertyVec)cdoc.ObjectFactory(nwEObjectType.eObjectType_nwOaPropertyVec, null, null);

                            // create a new Property (PropertyData)
                            InwOaProperty newProp = (InwOaProperty)cdoc.ObjectFactory(nwEObjectType.eObjectType_nwOaProperty, null, null);
                            newProp.name = dialog.PropertyName + "_InternalName";
                            newProp.UserName = dialog.PropertyName;
                            newProp.value = dialog.PropertyValue;
                            newcate.Properties().Add(newProp);

                            // Overwrite the existing property category with the newly created property category (existing + new)
                            cpropcates.SetUserDefined(1, dialog.CategoryName, dialog.CategoryName + "_InternalName", newcate);

                            //  return category;
                            break;
                        }
                    }

                    if (!categoryMatch)
                    {
                        // Category with the same name does not exist, add the new Category to item's CategoryDataCollection
                        // Create a new Property (PropertyData)
                        InwOaProperty newprop = (InwOaProperty)cdoc.ObjectFactory(nwEObjectType.eObjectType_nwOaProperty, null, null);
                        // Set PropertyName
                        newprop.name = dialog.PropertyName + "_InternalName";
                        // Set PropertyDisplayName  
                        newprop.UserName = dialog.PropertyName;
                        // Set PropertyValue
                        newprop.value = dialog.PropertyValue;
                        // Add PropertyData to Category
                        newcate.Properties().Add(newprop);
                        // Add CategoryData to item's CategoryDataCollection
                        cpropcates.SetUserDefined(0, dialog.CategoryName, dialog.CategoryName + "_InternalName", newcate);
                    }
                }
            }

            return 0;
        }
    }
}