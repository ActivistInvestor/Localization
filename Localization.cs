using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Resources;
using System.Text;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;

namespace Autodesk.AutoCAD.Runtime.Localization
{

   /// <summary>
   /// Work in progress...
   /// 
   /// Minimal documentation (so far)...
   /// 
   /// This class provides access to localized display names
   /// for the following elements:
   /// 
   ///   Managed types derived from Entity
   ///   
   ///   Properties of managed types derived from Entity
   ///   
   ///   Various Enum types
   ///   
   ///   Property Categories
   ///   
   /// You can investigate what display names can be returned
   /// by examining acdbmgd.resources.dll in the locale folder
   /// under the AutoCAD program directory. For example, on 
   /// english langauge AutoCAD that would be:
   /// 
   ///   Program Files\AutoCAD 20XX\en-US\acdbmgd.resources.dll
   ///    
   /// The contained resources can be opened and viewed using
   /// ILSpy, and probably other tools.
   /// 
   /// With ILSpy, open the dll and expand the Resources node,
   /// and you should find something like:
   /// 
   ///  Autodesk.AutoCAD.DatabaseServices.Entity.en-US.resources
   ///   
   /// Selecting it displays all of the resources in the right
   /// pane. The resource keys for Types, Properties, and Enums 
   /// have the prefixes "Type.", "Property.", and "Enum." that
   /// are followed by the type (or declaring type)'s full name.
   /// 
   /// The included code sorts all of that out and provides a 
   /// simple, high-level interface for accessing these values, 
   /// using a few extension methods.
   /// 
   /// </summary>

   public static class Localization
   {
      /// <summary>
      /// Due to the overhead involved in producing 
      /// results, we cache everything...
      /// </summary>

      static Dictionary<Assembly, ResourceManager> managers
         = new Dictionary<Assembly, ResourceManager>();

      static Dictionary<Type, string> types
         = new Dictionary<Type, string>();

      static Dictionary<PropertyInfo, string> properties
         = new Dictionary<PropertyInfo, string>();

      static Dictionary<Enum, string> enums
         = new Dictionary<Enum, string>();

      static LenientResourceManager manager = 
         new LenientResourceManager(typeof(Entity));

      /// <summary>
      /// The only resources currently supported are 
      /// AcDbMgd.resources.dll in the locale folder 
      /// ProgramFiles\AutoCAD 20??\<locale>\
      /// 
      /// These APIs were designed to fetch resources
      /// using the naming convention used in the above
      /// resources assembly (e.g., using the prefixes
      /// ("Class.", "Property.", and "Enum.").
      /// 
      /// On English AutoCAD the path to the file is:
      /// 
      ///   ProgramFiles\AutoCAD 20??\en-US\AcDbMgd.resources.dll
      ///   
      /// </summary>
      /// <param name="type"></param>
      /// <returns></returns>

      public static ResourceManager GetManager(Type type)
      {
         if(type == null)
            throw new ArgumentNullException(nameof(type));
         if(!managers.TryGetValue(type.Assembly, out ResourceManager value))
            managers[type.Assembly] = value = new LenientResourceManager(type);
         return value;
      }

      static Localization()
      {
         GetManager(typeof(Entity));
      }

      /// <summary>
      /// Manual fallback to find and load acdbmgd.resources.dll
      /// </summary>
      /// <param name="name"></param>
      /// <param name="type"></param>
      /// <returns></returns>
      /// <exception cref="ArgumentNullException"></exception>

      public static string GetLocalizedName(string name, Type type)
      {
         if(type == null)
            throw new ArgumentNullException(nameof(type));
         return GetLocalizedName(name, GetManager(type));
      }

      public static bool TryGetLocalizedName(string key, Type type, out string result)
      {
         if(string.IsNullOrWhiteSpace(key))
            throw new ArgumentException(nameof(key));
         if(type == null)
            throw new ArgumentNullException(nameof(type));
         return TryGetString(key, GetManager(type), out result);
      }

      public static string GetLocalizedName(string key, ResourceManager manager)
      {
         if(string.IsNullOrWhiteSpace(key))
            throw new ArgumentException(nameof(key));
         if(manager == null)
            throw new ArgumentNullException(nameof(manager));
         string str = manager.GetString(key);
         return string.IsNullOrWhiteSpace(str) ? key : str;
      }

      public static bool TryGetString(string key, ResourceManager manager, out string result)
      {
         if(string.IsNullOrWhiteSpace(key))
            throw new ArgumentException(nameof(key));
         if(manager == null)
            throw new ArgumentNullException(nameof(manager));
         string found = manager.GetString(key);
         bool flag = !string.IsNullOrWhiteSpace(found);
         result = flag ? found : string.Empty;
         return flag;
      }

      public static string GetLocalizedDisplayName(PropertyInfo property)
      {
         if(property == null)
            throw new ArgumentNullException(nameof(property));
         if(!properties.TryGetValue(property, out string result))
         {
            Type type = property.DeclaringType;
            string name = property.Name;
            if(!TryGetLocalizedName($"Property.{type.FullName}.{name}", type, out result))
            {
               DisplayNameAttribute dna = property.GetCustomAttribute<DisplayNameAttribute>();
               result = dna?.DisplayName ?? property.Name;
            }
            properties[property] = result;
         }
         return result;
      }

      /// <summary>
      /// Returns the localized display name of the given
      /// enum value.
      /// 
      /// The argument must be an enum value, for example:
      ///
      ///   (using Autodesk.AutoCAD.GraphicsInterface)
      ///   
      ///   GetLocalizedDisplayName(VisualStyleType.Wireframe2D);
      ///     
      /// </summary>
      /// <param name="value"></param>
      /// <returns></returns>

      public static string GetLocalizedDisplayName(Enum value)
      {
         if(!enums.TryGetValue(value, out string result))
         {
            Type enumType = value.GetType();
            string name = Enum.GetName(enumType, value);
            string key = $"Enum.{enumType.FullName}.{name}";
            if(!TryGetLocalizedName(key, enumType, out result))
            {
               FieldInfo fi = enumType.GetField(name);
               if(fi != null)
               {
                  var dna = fi.GetCustomAttribute<EnumDisplaNameAttribute>();
                  result = dna?.DisplayName ?? name;
               }
            }
            enums[value] = result;
         }
         return result;
      }

      /// <summary>
      /// Gets the localized display name of a property given
      /// its declaring type and property name
      /// 
      /// Note that the property name is case-sensitive.
      /// </summary>

      public static string GetLocalizedDisplayName(Type declaringType, string propertyName)
      {
         if(declaringType == null)
            throw new ArgumentNullException(nameof(declaringType));
         if(string.IsNullOrWhiteSpace(propertyName))
            throw new ArgumentException(nameof(propertyName));
         PropertyInfo pi = declaringType.GetProperty(propertyName);
         if(pi == null)
            throw new MissingMemberException(propertyName);
         return GetLocalizedDisplayName(pi);
      }

      /// <summary>
      /// Gets the localized display name for the given type.
      /// </summary>

      public static string GetLocalizedDisplayName(Type type)
      {
         if(type == null)
            throw new ArgumentNullException(nameof(type));
         string result;
         if(!types.TryGetValue(type, out result))
         {
            if(!TryGetLocalizedName($"Class.{type.FullName}", type, out result))
            {
               var dna = type.GetCustomAttribute<DisplayNameAttribute>();
               result = dna?.DisplayName ?? type.Name;
            }
            types[type] = result;
         }
         return result;
      }

      /// <summary>
      /// Inserts the given sequence of delimiting characters 
      /// between each occurrence of an uppercase character that
      /// immediately follows a lowercase character, excluding 
      /// the first character in the sequence.
      /// 
      /// For example, "MoeLarryCurly" => "Moe Larry Curly";
      /// 
      /// Currently not used, but may be applied to the result
      /// of display name requests for enum fields, when there
      /// is no resource string found and the method falls back 
      /// to the enum member's declared name.
      /// 
      /// </summary>

      static string FixUpName(string name, string delimiter = " ")
      {
         if(name == null)
            throw new ArgumentNullException(nameof(name));
         if(name.Length < 2)
            return name;
         StringBuilder sb = new StringBuilder(name[0]);
         bool flag = char.IsUpper(name[0]);   // preserves camel-case
         for(int i = 1; i < name.Length; i++)
         {
            char c = name[i];
            flag = char.IsUpper(c) && !flag;
            if(flag)
               sb.Append(delimiter + c);
            else
               sb.Append(c);
         }
         return sb.ToString();
      }

      /// <summary>
      /// Extension method targeting System.Type, searches 
      /// for and returns a localized display name for the 
      /// type, if defined in the containing assembly's
      /// resources.
      /// 
      /// The returned string is primarily intended for 
      /// display in the UI, in preference to the class's 
      /// name or a runtime class' DxfName.
      /// 
      /// This method may only be applicable for types from 
      /// AcDbMgd.dll, and has not been tested against types 
      /// from other assemblies. The lookup key used to find
      /// the resource string is $"Class.{type.Fullname}".
      /// 
      /// If no localized name is found, the method looks for
      /// a DisplaNameAttribute applied to the type, and if 
      /// found, returns the value of its DisplayName property.
      /// Otherwise, the value of the type's Name property is
      /// is returned.
      /// 
      /// Note that the value returned by this method is not
      /// merely a local translation of the type name, it is 
      /// the UI display name for the type (which means that 
      /// even in English language versions it will return a
      /// value that is more-appropriate for display in a UI).
      /// 
      /// Examples of where you will find display names used
      /// in place of Type names include AutoCAD's Properties 
      /// palette, and AutoCAD's Data Extraction UI.
      /// </summary>

      public static string GetDisplayName(this Type type)
      {
         if(type == null)
            throw new ArgumentNullException(nameof(type));
         return GetLocalizedDisplayName(type);
      }

      /// <summary>
      /// Extension method targeting PropertyInfo
      /// 
      /// Gets the localized display name for a declared
      /// property, given a PropertyInfo representing the 
      /// property.
      /// 
      /// If no localized display name is found in the
      /// declaring type's containing assembly resources,
      /// this method will then try to get the value of a
      /// DisplayNameAttribute applied to the property, and
      /// return same if found. If no DisplayNameAttribute 
      /// is found, the Name of the property is returned.
      /// 
      /// GetPropertyDisplayName() can also be used to get 
      /// the localized display name of a property, given 
      /// the declaring type and property name.
      /// </summary>

      public static string GetDisplayName(this PropertyInfo property)
      {
         if(property == null)
            throw new ArgumentNullException(nameof(property));
         return GetLocalizedDisplayName(property);
      }

      /// <summary>
      /// Extension method targeting subtypes of DBObject:
      /// 
      /// Given an instance of a type that is a subclass of
      /// DBObject, and a string that identifies a property
      /// of the type, this method returns the display name
      /// of the property.
      /// </summary>

      public static string GetPropertyDisplayName(this Type type, string propertyName)
      {
         if(type == null)
            throw new ArgumentNullException(nameof(type));
         if(string.IsNullOrWhiteSpace(propertyName))
            throw new ArgumentException(nameof(propertyName));
         return GetLocalizedDisplayName(type, propertyName);
      }

      /// <summary>
      /// Gets the localized display name for an enum value.
      /// 
      /// If no localized display name can be found in the
      /// enum type's assembly resources, this method returns
      /// the enum's name.
      /// 
      /// </summary>
      /// <param name="value"></param>
      /// <returns></returns>
      /// <exception cref="ArgumentNullException"></exception>

      public static string GetDisplayName(this Enum value)
      {
         if(value == null)
            throw new ArgumentNullException(nameof(value));
         return GetLocalizedDisplayName(value);
      }

   }

   /// <summary>
   /// The framework's DisplayNameAttribute cannot be
   /// applied to enum fields. This derivative can, and
   /// methods of the Localization class check for it.
   /// </summary>

   [AttributeUsage(AttributeTargets.Field, AllowMultiple = false)]
   public class EnumDisplaNameAttribute : DisplayNameAttribute
   {
      public EnumDisplaNameAttribute(string displaName)
         : base(displaName)
      {
         if(string.IsNullOrWhiteSpace(displaName))
            throw new ArgumentException(nameof(EnumDisplaNameAttribute));
      }
   }
}

namespace Test
{
   using Autodesk.AutoCAD.Runtime.Localization;
   using Autodesk.AutoCAD.GraphicsInterface;

   [Flags]
   public enum MyEnum
   {
      [EnumDisplaName("Field 1")]
      Field1 = 1,
      Field2 = 2,
      Field3 = 3
   }

   public static class TestLocalizedNameHelper
   {
      [CommandMethod("TESTLOCALIZEDNAMEHELPER")]
      public static void Test()
      {

         VisualStyleType styleType = VisualStyleType.FlatWithEdges;
         string vsDisplaName = styleType.GetDisplayName();
         Write($"VisualStyleType.FlatWithEdges.DisplayName() = {vsDisplaName}");

         PropertyInfo layerProp = typeof(Entity).GetProperty("Layer");

         string localPolyLineName = typeof(Polyline).GetDisplayName();
         Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(
            $"Localized name of Polyline: {localPolyLineName}");
      }

      static void Write(string msg, params object[] args)
      {

      }
   }
}
