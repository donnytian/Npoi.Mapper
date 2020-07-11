using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Reflection.Emit;
using System.Security.Cryptography;
using System.Text;

namespace Npoi.Mapper
{
    /// <summary>
    /// Methods for creating anonymous type dynamically.
    /// </summary>
    public static class AnonymousTypeFactory
    {
        private static readonly ModuleBuilder ModuleBuilder;
        private static readonly object SyncRoot = new object();

        static AnonymousTypeFactory()
        {
            var assemblyName = new AssemblyName { Name = "MyAnonymousTypes" };

#if NET45 || NET40
            var assemblyBuilder = AppDomain.CurrentDomain.DefineDynamicAssembly(assemblyName, AssemblyBuilderAccess.Run);
            ModuleBuilder = assemblyBuilder.DefineDynamicModule(assemblyName.Name);
#else
            var assemblyBuilder = AssemblyBuilder.DefineDynamicAssembly(assemblyName, AssemblyBuilderAccess.Run);
            ModuleBuilder = assemblyBuilder.DefineDynamicModule(assemblyName.Name);
#endif

        }

        // This version of the static constructor can be used temporarily to save the dynamic DLL to get a copy of the generated IL.
        //public static readonly AssemblyBuilder assemblyBuilder;
        //static AnonymousTypeFactory()
        //{
        //    var assemblyName = new AssemblyName { Name = "MyAnonymousTypes" };
        //    assemblyBuilder = System.Threading.Thread.GetDomain().DefineDynamicAssembly(assemblyName, AssemblyBuilderAccess.RunAndSave, @"C:\Temp\");
        //    ModuleBuilder = assemblyBuilder.DefineDynamicModule(assemblyName.Name, "anonymousTypes.dll");

        //    // This would need to be called from somewhere else after a dynamic anonymous type is created
        //    // AnonymousTypeFactory.ModuleBuilder.Save("MyAnonymousTypes.dll");
        //}

        /// <summary>
        /// Creates an object from an anonymous type with the specified name/value pairs.
        /// </summary>
        /// <param name="valueDictionary">mappings of property names to values</param>
        /// <returns>the object</returns>
        public static object CreateObject(IDictionary<string, object> valueDictionary)
        {
            var typeDictionary = valueDictionary.ToDictionary(kv => kv.Key, kv => kv.Value != null ? kv.Value.GetType() : typeof(object));
            var anonymousType = CreateType(typeDictionary);
            return CreateObject(valueDictionary, anonymousType);
        }

        /// <summary>
        /// Creates an object for the specified name/value pairs with the specified anonymous type.
        /// </summary>
        /// <remarks>
        /// This version of the <c>CreateObject</c> method can be used when you need more control 
        /// over how the anonymous type is generated.
        /// </remarks>
        /// <param name="valueDictionary">mappings of property names to values</param>
        /// <param name="anonymousType">the anonymous type to use for the resulting object</param>
        /// <returns>the object</returns>
        public static object CreateObject(IDictionary<string, object> valueDictionary, Type anonymousType)
        {
            var constructor = anonymousType.GetConstructors().Single();
            var parameters = constructor.GetParameters();
            var parameterValues = parameters.Select(p => valueDictionary[p.Name]).ToArray();
            return constructor.Invoke(parameterValues);
        }

        /// <summary>
        /// Creates an anonymous type for the specified property name/type pairs.
        /// </summary>
        /// <param name="typePairs">mappings of property names to types</param>
        /// <returns>the anonymous type</returns>
        public static Type CreateType(IEnumerable<KeyValuePair<string, Type>> typePairs)
        {
            return CreateType(typePairs, isMutable: false, parent: null);
        }

        /// <summary>
        /// Creates an anonymous type that may be mutable for the specified property name/type pairs.
        /// </summary>
        /// <param name="typePairs">mappings of property names to types</param>
        /// <param name="isMutable">true if the anonymous type's properties should have setters</param>
        /// <returns>the anonymous type</returns>
        public static Type CreateType(IEnumerable<KeyValuePair<string, Type>> typePairs, bool isMutable)
        {
            return CreateType(typePairs, isMutable: isMutable, parent: null);
        }

        /// <summary>
        /// Creates an anonymous type that may be mutable for the specified property name/type pairs with the (optional) parent type.
        /// </summary>
        /// <param name="typePairs">mappings of property names to types</param>
        /// <param name="isMutable">true if the anonymous type's properties should have setters</param>
        /// <param name="parent">the parent of the anonymous type, or null to not specify a base type</param>
        /// <returns>the anonymous type</returns>
        public static Type CreateType(IEnumerable<KeyValuePair<string, Type>> typePairs, bool isMutable, Type parent)
        {
            if (typePairs == null) throw new ArgumentNullException(nameof(typePairs));

            var keyValuePairs = typePairs as KeyValuePair<string, Type>[] ?? typePairs.ToArray();
            var propertyNames = keyValuePairs.Select(pair => pair.Key);
            var genericTypeDefinition = GetOrCreateGenericTypeDefinition(propertyNames.ToList(), isMutable: isMutable, parent: parent);

            var propertyTypes = keyValuePairs.Select(pair => pair.Value);
            return genericTypeDefinition.MakeGenericType(propertyTypes.ToArray());
        }

        /// <summary>
        /// Creates an anonymous generic type definition for the specified property names.
        /// </summary>
        /// <param name="propertyNames">the property names</param>
        /// <returns>the anonymous generic type definition</returns>
        public static Type CreateGenericTypeDefinition(IEnumerable<string> propertyNames)
        {
            return CreateGenericTypeDefinition(propertyNames, isMutable: false, parent: null);
        }

        /// <summary>
        /// Creates an anonymous generic type definition that may be mutable with the specified property names.
        /// </summary>
        /// <param name="propertyNames">the property names</param>
        /// <param name="isMutable">true if the properties should have setters</param>
        /// <returns>the anonymous generic type definition</returns>
        public static Type CreateGenericTypeDefinition(IEnumerable<string> propertyNames, bool isMutable)
        {
            return CreateGenericTypeDefinition(propertyNames, isMutable: isMutable, parent: null);
        }

        /// <summary>
        /// Creates an anonymous generic type definition that may be mutable for the specified property names and (optional) parent type.
        /// </summary>
        /// <param name="propertyNames">the property names</param>
        /// <param name="isMutable">true if the properties should have setters</param>
        /// <param name="parent">the parent class of the anonymous type, or null if it has no parent class</param>
        /// <returns>the anonymous generic type definition</returns>
        public static Type CreateGenericTypeDefinition(IEnumerable<string> propertyNames, bool isMutable, Type parent)
        {
            if (propertyNames == null) throw new ArgumentNullException(nameof(propertyNames));
            return GetOrCreateGenericTypeDefinition(propertyNames.ToList(), isMutable: isMutable, parent: parent);
        }

        private static Type GetOrCreateGenericTypeDefinition(ICollection<string> propertyNames, bool isMutable, Type parent)
        {
            if (!propertyNames.Any())
            {
                throw new ArgumentOutOfRangeException(nameof(propertyNames), propertyNames.Count, "At least one property name is required to create an anonymous type");
            }
            if (parent != null && !parent.GetConstructors().Any(c => !c.GetParameters().Any()))
            {
                throw new ArgumentException(
                    $"Parent type \"{parent.FullName}\" is not supported because it does not have a default constructor");
            }

            var genericTypeDefinitionName = GenerateGenericTypeDefinitionName(
                propertyNames,
                isMutable: isMutable,
                parent: parent
            );

            // We need to check for the type and define/create it as one atomic operation, 
            // otherwise we could get a TypeBuilder back instead of a full Type.
            Type genericTypeDefinition;
            lock (SyncRoot)
            {
                genericTypeDefinition = ModuleBuilder.GetType(genericTypeDefinitionName);
                if (genericTypeDefinition == null)
                {
                    genericTypeDefinition = CreateGenericTypeDefinitionNoLock(
                        genericTypeDefinitionName,
                        propertyNames,
                        isMutable: isMutable,
                        parent: parent
                    );
                }
            }
            return genericTypeDefinition;
        }

        private static string GenerateGenericTypeDefinitionName(ICollection<string> propertyNames, bool isMutable, Type parent)
        {
            // A real anonymous type is named something like "<>f__AnonymousType0`2" (for the first anonymous type generated with two properties).

            // We'll mostly try to match this format, but with a couple differences:
            // * Add our library name (to avoid any possible name conflict with real anonymous types)
            // * Use a hash of the options (property names, mutable, parent type) instead of the counter (the "0" in the above example)

            // That counter in a real anonymous type is zero for the first anonymous type, one for the second, two for the third, etc.
            // We could store a global counter to do the same thing, but then using the same set of property names 
            // could generate a different type name depending on how many types you've already generated.
            // I'd rather have a deterministic name generation strategy, so we're going to use a hash of the 
            // key values instead of a global counter.

            // It doesn't matter exactly this key string is formatted or how long it is since we'll just be computing a hash from it. 
            // So I guess we'll just use a JSON format for now?
            var keyJsonBuilder = new StringBuilder();
            keyJsonBuilder.Append('{');
            keyJsonBuilder.Append("properties=[");
            keyJsonBuilder.Append(string.Join(",", propertyNames.Select(n => '"' + n.Replace("\"", "\"\"") + '"')));
            keyJsonBuilder.Append(']');
            if (isMutable)
            {
                keyJsonBuilder.Append(",isMutable=true");
            }
            if (parent != null)
            {
                keyJsonBuilder.Append(",parent=\"");
                keyJsonBuilder.Append(parent.FullName);
                keyJsonBuilder.Append("\"");
            }
            // If we add support for interfaces, then we'll need them included in the key string too...
            //if (interfaces != null && interfaces.Any())
            //{
            //    keyJsonBuilder.Append(",interfaces=[");
            //    keyJsonBuilder.Append(string.Join(",", interfaces.Select(x => x.FullName).OrderBy(n => n).Select(n => '"' + n.Replace("\"", "\"\"") + '"')));
            //    keyJsonBuilder.Append("]");
            //}
            keyJsonBuilder.Append('}');

            string keyHashHexString;
            using (var hasher = new SHA1CryptoServiceProvider())
            {
                var hashBytes = hasher.ComputeHash(Encoding.UTF8.GetBytes(keyJsonBuilder.ToString()));
                keyHashHexString = BitConverter.ToString(hashBytes).Replace("-", string.Empty);
            }

            string genericTypeDefinitionName = $"<>f__MyAnonymousType{keyHashHexString}`{propertyNames.Count}";
            return genericTypeDefinitionName;
        }

        private static Type CreateGenericTypeDefinitionNoLock(string genericTypeDefinitionName, ICollection<string> propertyNames, bool isMutable, Type parent)
        {
            var typeBuilder = ModuleBuilder.DefineType(genericTypeDefinitionName,
                attr: TypeAttributes.Public | TypeAttributes.AutoLayout
                | TypeAttributes.AnsiClass | TypeAttributes.Class
                | TypeAttributes.Sealed | TypeAttributes.BeforeFieldInit,
                parent: parent
            );
            var typeParameterNames = propertyNames
                .Select(propertyName => $"<{propertyName}>j__TPar")
                .ToArray();
            var typeParameters = typeBuilder.DefineGenericParameters(typeParameterNames);

            var typeParameterPairs = propertyNames.Zip(typeParameters,
                (propertyName, typeParameter) => new KeyValuePair<string, GenericTypeParameterBuilder>(propertyName, typeParameter)
            ).ToArray();

            var fieldBuilders = new List<FieldBuilder>(typeParameterPairs.Length);
            foreach (var pair in typeParameterPairs)
            {
                var propertyName = pair.Key;
                var typeParameter = pair.Value;
                var fieldAttributes = FieldAttributes.Private;
                if (!isMutable)
                {
                    fieldAttributes = fieldAttributes | FieldAttributes.InitOnly;
                }
                var fieldBuilder = typeBuilder.DefineField($"<{propertyName}>i__Field", typeParameter, fieldAttributes);
                fieldBuilders.Add(fieldBuilder);
                var property = typeBuilder.DefineProperty(propertyName, PropertyAttributes.None, typeParameter, Type.EmptyTypes);

                var getMethodBuilder = typeBuilder.DefineMethod(
                    name: $"get_{propertyName}",
                    attributes: MethodAttributes.PrivateScope | MethodAttributes.Public | MethodAttributes.HideBySig | MethodAttributes.SpecialName,
                    callingConvention: CallingConventions.Standard | CallingConventions.HasThis,
                    returnType: typeParameter,
                    parameterTypes: Type.EmptyTypes
                );
                var getMethodIlGenerator = getMethodBuilder.GetILGenerator();
                getMethodIlGenerator.Emit(OpCodes.Ldarg_0);
                getMethodIlGenerator.Emit(OpCodes.Ldfld, fieldBuilder);
                getMethodIlGenerator.Emit(OpCodes.Ret);
                property.SetGetMethod(getMethodBuilder);

                if (isMutable)
                {
                    var setMethodBuilder = typeBuilder.DefineMethod(
                        name: $"set_{propertyName}",
                        attributes: MethodAttributes.PrivateScope | MethodAttributes.Public | MethodAttributes.HideBySig | MethodAttributes.SpecialName,
                        callingConvention: CallingConventions.Standard | CallingConventions.HasThis,
                        returnType: null,
                        parameterTypes: new Type[] { typeParameter }
                    );
                    var setMethodIlGenerator = setMethodBuilder.GetILGenerator();
                    setMethodIlGenerator.Emit(OpCodes.Ldarg_0);
                    setMethodIlGenerator.Emit(OpCodes.Ldarg_1);
                    setMethodIlGenerator.Emit(OpCodes.Stfld, fieldBuilder);
                    setMethodIlGenerator.Emit(OpCodes.Ret);
                    property.SetSetMethod(setMethodBuilder);
                }
            }

            var defaultConstructor = parent?.GetConstructors().FirstOrDefault(c => !c.GetParameters().Any());
            DefineDefaultConstructor(typeBuilder, defaultConstructor);
            // We don't need this for now, we only need a parameterless constructor.
            //DefineConstructor(typeBuilder, typeParameters.ToArray<Type>(), propertyNames, fieldBuilders, defaultConstructor);

            DefineEqualsMethod(typeBuilder, fieldBuilders);
            DefineGetHashCodeMethod(typeBuilder, fieldBuilders);

            var fieldPairs = propertyNames.Zip(fieldBuilders,
                (propertyName, fieldBuilder) => new KeyValuePair<string, FieldBuilder>(propertyName, fieldBuilder)
            ).ToArray();
            DefineToStringMethod(typeBuilder, fieldPairs);

#if NET40
            return typeBuilder.CreateType();
#else
            return typeBuilder.CreateTypeInfo();
#endif
        }

        private static void DefineDefaultConstructor(TypeBuilder typeBuilder, ConstructorInfo baseConstructor = null)
        {
            var constructorBuilder = typeBuilder.DefineConstructor(
                attributes: MethodAttributes.PrivateScope | MethodAttributes.Public | MethodAttributes.HideBySig | MethodAttributes.SpecialName | MethodAttributes.RTSpecialName,
                callingConvention: CallingConventions.Standard | CallingConventions.HasThis,
                parameterTypes: Type.EmptyTypes
            );

            var constructorIlGenerator = constructorBuilder.GetILGenerator();
            constructorIlGenerator.Emit(OpCodes.Ldarg_0);
            constructorIlGenerator.Emit(OpCodes.Call, typeof(object).GetConstructors().Single());

            if (baseConstructor != null)
            {
                constructorIlGenerator.Emit(OpCodes.Ldarg_0);
                constructorIlGenerator.Emit(OpCodes.Call, baseConstructor);
            }
            constructorIlGenerator.Emit(OpCodes.Ret);
        }

        // ReSharper disable once UnusedMember.Local
        private static void DefineConstructor(
            TypeBuilder typeBuilder,
            Type[] typeParameters,
            ICollection<string> propertyNames,
            ICollection<FieldBuilder> fieldBuilders,
            ConstructorInfo baseConstructor = null)
        {
            var constructorBuilder = typeBuilder.DefineConstructor(
                attributes: MethodAttributes.PrivateScope | MethodAttributes.Public | MethodAttributes.HideBySig | MethodAttributes.SpecialName | MethodAttributes.RTSpecialName,
                callingConvention: CallingConventions.Standard | CallingConventions.HasThis,
                parameterTypes: typeParameters
            );
            foreach (var o in propertyNames.Select((propertyName, index) => new { propertyName, index }))
            {
                constructorBuilder.DefineParameter(o.index + 1, ParameterAttributes.None, o.propertyName);
            }

            var constructorIlGenerator = constructorBuilder.GetILGenerator();
            constructorIlGenerator.Emit(OpCodes.Ldarg_0);
            constructorIlGenerator.Emit(OpCodes.Call, typeof(object).GetConstructors().Single());
            foreach (var obj in fieldBuilders.Select((fieldBuilder, index) => new { fieldBuilder, index }))
            {
                constructorIlGenerator.Emit(OpCodes.Ldarg_0);

                var field = obj.fieldBuilder;
                var index = obj.index;
                switch (index)
                {
                    case 0:
                        constructorIlGenerator.Emit(OpCodes.Ldarg_1);
                        break;
                    case 1:
                        constructorIlGenerator.Emit(OpCodes.Ldarg_2);
                        break;
                    case 2:
                        constructorIlGenerator.Emit(OpCodes.Ldarg_3);
                        break;
                    default:
                        constructorIlGenerator.Emit(OpCodes.Ldarg_S, index + 1);
                        break;
                }
                constructorIlGenerator.Emit(OpCodes.Stfld, field);
            }
            if (baseConstructor != null)
            {
                constructorIlGenerator.Emit(OpCodes.Ldarg_0);
                constructorIlGenerator.Emit(OpCodes.Call, baseConstructor);
            }
            constructorIlGenerator.Emit(OpCodes.Ret);
        }

        private static void DefineEqualsMethod(TypeBuilder typeBuilder, ICollection<FieldBuilder> fields)
        {
            var equalsMethodBuilder = typeBuilder.DefineMethod(
                name: "Equals",
                attributes: MethodAttributes.Public | MethodAttributes.HideBySig |
                    MethodAttributes.Virtual | MethodAttributes.Final,
                returnType: typeof(bool),
                parameterTypes: new[] { typeof(object) }
            );
            equalsMethodBuilder.DefineParameter(1, ParameterAttributes.None, "value");

            var il = equalsMethodBuilder.GetILGenerator();

            il.DeclareLocal(typeBuilder);
            il.DeclareLocal(typeof(bool));

            var label1 = il.DefineLabel();
            var label2 = il.DefineLabel();
            var label3 = il.DefineLabel();

            il.Emit(OpCodes.Ldarg_1);
            il.Emit(OpCodes.Isinst, typeBuilder);
            il.Emit(OpCodes.Stloc_0);
            il.Emit(OpCodes.Ldloc_0);

            // Only the last five fields can use the short form of the branch.
            const int maximumShortBranchFieldCount = 5;
            var shortBranchThreshold = Math.Max(fields.Count - maximumShortBranchFieldCount, 0);

            int currentFieldIndex = 0;
            foreach (var field in fields)
            {
                var equalityComparerGenericTypeDefinition = typeof(EqualityComparer<>);
                var equalityComparerEqualsGenericMethodDefinition = equalityComparerGenericTypeDefinition
                    .GetMethods().Single(m => m.Name == "Equals" && m.GetParameters().Length == 2);
                var equalityComparerDefaultGenericPropertyGetterDefinition = equalityComparerGenericTypeDefinition
                    .GetProperty("Default", BindingFlags.Public | BindingFlags.Static)
                    .GetGetMethod();

                var equalityComparerType = equalityComparerGenericTypeDefinition.MakeGenericType(field.FieldType);
                var equalityComparerEqualsMethod = TypeBuilder.GetMethod(equalityComparerType, equalityComparerEqualsGenericMethodDefinition);
                var equalityComparerDefaultPropertyGetter = TypeBuilder.GetMethod(equalityComparerType, equalityComparerDefaultGenericPropertyGetterDefinition);

                if (currentFieldIndex >= shortBranchThreshold)
                {
                    il.Emit(OpCodes.Brfalse_S, label1);
                }
                else
                {
                    il.Emit(OpCodes.Brfalse, label1);
                }
                il.EmitCall(OpCodes.Call, equalityComparerDefaultPropertyGetter, null);

                il.Emit(OpCodes.Ldarg_0);
                il.Emit(OpCodes.Ldfld, field);

                il.Emit(OpCodes.Ldloc_0);
                il.Emit(OpCodes.Ldfld, field);

                il.EmitCall(OpCodes.Callvirt, equalityComparerEqualsMethod, null);

                currentFieldIndex++;
            }
            il.Emit(OpCodes.Br_S, label2);

            il.MarkLabel(label1);
            il.Emit(OpCodes.Ldc_I4_0);
            il.MarkLabel(label2);
            il.Emit(OpCodes.Nop);
            il.Emit(OpCodes.Stloc_1);
            il.Emit(OpCodes.Br_S, label3);

            il.MarkLabel(label3);
            il.Emit(OpCodes.Ldloc_1);
            il.Emit(OpCodes.Ret);

            typeBuilder.DefineMethodOverride(equalsMethodBuilder, typeof(object).GetMethod("Equals", new[] { typeof(object) }));
        }

        private static void DefineGetHashCodeMethod(TypeBuilder typeBuilder, IEnumerable<FieldBuilder> fields)
        {
            var getHashCodeMethodBuilder = typeBuilder.DefineMethod(
                name: "GetHashCode",
                attributes: MethodAttributes.Public | MethodAttributes.HideBySig |
                    MethodAttributes.Virtual | MethodAttributes.Final,
                returnType: typeof(int),
                parameterTypes: Type.EmptyTypes
            );

            var il = getHashCodeMethodBuilder.GetILGenerator();

            il.DeclareLocal(typeof(int));
            il.DeclareLocal(typeof(int));

            int hashSeed = 0;
            const int hashMultiplier = -1521134295;
            var fieldBuilders = fields as FieldBuilder[] ?? fields.ToArray();
            foreach (var field in fieldBuilders)
            {
                unchecked
                {
                    hashSeed = (hashSeed * hashMultiplier) + field.Name.GetHashCode();
                }
            }

            il.Emit(OpCodes.Ldc_I4, hashSeed);
            il.Emit(OpCodes.Stloc_0);

            foreach (var field in fieldBuilders)
            {
                var equalityComparerGenericTypeDefinition = typeof(EqualityComparer<>);
                var equalityComparerDefaultGenericPropertyGetterDefinition = equalityComparerGenericTypeDefinition
                    .GetProperty("Default", BindingFlags.Public | BindingFlags.Static)
                    .GetGetMethod();
                var equalityComparerGetHashCodeGenericMethodDefinition = equalityComparerGenericTypeDefinition
                    .GetMethods().Single(m => m.Name == "GetHashCode" && m.GetParameters().Length == 1);

                var equalityComparerType = equalityComparerGenericTypeDefinition.MakeGenericType(field.FieldType);
                var equalityComparerDefaultPropertyGetter = TypeBuilder.GetMethod(equalityComparerType, equalityComparerDefaultGenericPropertyGetterDefinition);
                var equalityComparerGetHashCodeMethod = TypeBuilder.GetMethod(equalityComparerType, equalityComparerGetHashCodeGenericMethodDefinition);

                il.Emit(OpCodes.Ldc_I4, hashMultiplier);
                il.Emit(OpCodes.Ldloc_0);
                il.Emit(OpCodes.Mul);

                il.EmitCall(OpCodes.Call, equalityComparerDefaultPropertyGetter, null);

                il.Emit(OpCodes.Ldarg_0);
                il.Emit(OpCodes.Ldfld, field);
                il.EmitCall(OpCodes.Callvirt, equalityComparerGetHashCodeMethod, null);
                il.Emit(OpCodes.Add);
                il.Emit(OpCodes.Stloc_0);
            }

            il.Emit(OpCodes.Ldloc_0);
            il.Emit(OpCodes.Stloc_1);
            il.Emit(OpCodes.Ldloc_1);
            il.Emit(OpCodes.Ret);

            typeBuilder.DefineMethodOverride(getHashCodeMethodBuilder, typeof(object).GetMethod("GetHashCode"));
        }

        private static void DefineToStringMethod(TypeBuilder typeBuilder, IEnumerable<KeyValuePair<string, FieldBuilder>> fieldPairs)
        {
            var toStringMethodBuilder = typeBuilder.DefineMethod(
                name: "ToString",
                attributes: MethodAttributes.Public | MethodAttributes.HideBySig |
                    MethodAttributes.Virtual | MethodAttributes.Final,
                returnType: typeof(string),
                parameterTypes: Type.EmptyTypes
            );

            var il = toStringMethodBuilder.GetILGenerator();

            il.DeclareLocal(typeof(StringBuilder));
            il.DeclareLocal(typeof(string));

            // ReSharper disable once AssignNullToNotNullAttribute
            il.Emit(OpCodes.Newobj, typeof(StringBuilder).GetConstructor(Type.EmptyTypes));
            il.Emit(OpCodes.Stloc_0);

            var appendStringMethod = typeof(StringBuilder).GetMethod("Append", new[] { typeof(string) });
            var appendObjectMethod = typeof(StringBuilder).GetMethod("Append", new[] { typeof(object) });

            bool isFirst = true;
            foreach (var pair in fieldPairs)
            {
                var propertyName = pair.Key;
                var field = pair.Value;

                var sb = new StringBuilder();
                if (isFirst)
                {
                    sb.Append("{ ");
                    isFirst = false;
                }
                else
                {
                    sb.Append(", ");
                }

                sb.Append(propertyName);
                sb.Append(" = ");

                il.Emit(OpCodes.Ldloc_0);
                il.Emit(OpCodes.Ldstr, sb.ToString());
                il.Emit(OpCodes.Callvirt, appendStringMethod);
                il.Emit(OpCodes.Pop);

                il.Emit(OpCodes.Ldloc_0);
                il.Emit(OpCodes.Ldarg_0);
                il.Emit(OpCodes.Ldfld, field);
                il.Emit(OpCodes.Box, field.FieldType);
                il.Emit(OpCodes.Callvirt, appendObjectMethod);
                il.Emit(OpCodes.Pop);
            }

            il.Emit(OpCodes.Ldloc_0);
            il.Emit(OpCodes.Ldstr, " }");
            il.Emit(OpCodes.Callvirt, appendStringMethod);
            il.Emit(OpCodes.Pop);

            il.Emit(OpCodes.Ldloc_0);
            il.Emit(OpCodes.Callvirt, typeof(object).GetMethod("ToString"));
            il.Emit(OpCodes.Stloc_1);
            il.Emit(OpCodes.Ldloc_1);
            il.Emit(OpCodes.Ret);

            typeBuilder.DefineMethodOverride(toStringMethodBuilder, typeof(object).GetMethod("ToString"));
        }
    }
}
