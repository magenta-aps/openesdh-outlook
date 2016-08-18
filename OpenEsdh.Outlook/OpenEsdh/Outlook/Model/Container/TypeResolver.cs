namespace OpenEsdh.Outlook.Model.Container
{
    using OpenEsdh.Outlook.Model.Logging;
    using System;
    using System.Collections.Generic;
    using System.Runtime.CompilerServices;

    public abstract class TypeResolver : IDisposable
    {
        private static TypeResolver _current = null;
        private static object _lock = new object();
        protected readonly Dictionary<Type, object> _singletons = new Dictionary<Type, object>();
        private readonly Dictionary<Type, CreateCode> _typeToCreateCode = new Dictionary<Type, CreateCode>();
        private readonly Dictionary<Type, CreateCodeWithParam> _typeToCreateWithParamCode = new Dictionary<Type, CreateCodeWithParam>();

        public TypeResolver()
        {
            this.BuildComponents();
        }

        public void AddComponent<T>(CreateCode CreateCode)
        {
            Logger.Current.LogInformation("Adding type " + typeof(T).Name, "");
            if (this._typeToCreateCode.ContainsKey(typeof(T)))
            {
                this._typeToCreateCode.Remove(typeof(T));
            }
            this._typeToCreateCode.Add(typeof(T), CreateCode);
        }

        public void AddComponentWithParam<T>(CreateCodeWithParam CreateCode)
        {
            Logger.Current.LogInformation("Adding type " + typeof(T).Name, "");
            if (this._typeToCreateWithParamCode.ContainsKey(typeof(T)))
            {
                this._typeToCreateWithParamCode.Remove(typeof(T));
            }
            this._typeToCreateWithParamCode.Add(typeof(T), CreateCode);
        }

        protected abstract void BuildComponents();
        public T Create<T>()
        {
            T local;
            try
            {
                local = (T) this._typeToCreateCode[typeof(T)]();
            }
            catch (Exception exception)
            {
                Logger.Current.LogException(exception, "");
                Logger.Current.LogWarning(typeof(T).Name + " was not found", "");
                throw exception;
            }
            return local;
        }

        public T Create<T>(object Param)
        {
            T local;
            try
            {
                local = (T) this._typeToCreateWithParamCode[typeof(T)](Param);
            }
            catch (Exception exception)
            {
                Logger.Current.LogException(exception, "");
                Logger.Current.LogWarning(typeof(T).Name + " was not found", "");
                throw exception;
            }
            return local;
        }

        public virtual void Dispose()
        {
        }

        public void Replace<T>(CreateCode CreateCode)
        {
            try
            {
                this._typeToCreateCode[typeof(T)] = CreateCode;
            }
            catch (Exception exception)
            {
                Logger.Current.LogException(exception, "");
                Logger.Current.LogWarning(typeof(T).Name + " was not found", "");
                throw exception;
            }
        }

        public static TypeResolver Current
        {
            get
            {
                return _current;
            }
            set
            {
                lock (_lock)
                {
                    _current = value;
                }
            }
        }

        public Dictionary<Type, object> Singletons
        {
            get
            {
                return this._singletons;
            }
        }

        public delegate object CreateCode();

        public delegate object CreateCodeWithParam(object param);
    }
}

