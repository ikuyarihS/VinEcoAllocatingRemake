#region

using System;
using System.Collections.Concurrent;

#endregion

namespace VinEcoAllocatingRemake.AllocatingInventory
{
    #region

    #endregion

    /// <summary>
    ///     The object pool.
    /// </summary>
    /// <typeparam name="T">
    ///     Type in pool.
    /// </typeparam>
    public class ObjectPool<T>
    {
        /// <summary>
        ///     The object generator.
        /// </summary>
        private readonly Func<T> objectGenerator;

        /// <summary>
        ///     The objects.
        /// </summary>
        private readonly ConcurrentBag<T> objects;

        /// <summary>
        ///     Initializes a new instance of the <see cref="ObjectPool{T}" /> class.
        /// </summary>
        /// <param name="objectGenerator">
        ///     The object generator.
        /// </param>
        public ObjectPool(Func<T> objectGenerator)
        {
            objects = new ConcurrentBag<T>();
            this.objectGenerator = objectGenerator ?? throw new ArgumentNullException(nameof(objectGenerator));
        }

        /// <summary>
        ///     The get object.
        /// </summary>
        /// <returns>
        ///     The <see cref="T" />.
        /// </returns>
        public T GetObject()
        {
            return objects.TryTake(out T item)
                ? item
                : objectGenerator();
        }

        /// <summary>
        ///     The put object.
        /// </summary>
        /// <param name="item">
        ///     The item.
        /// </param>
        public void PutObject(T item)
        {
            objects.Add(item);
        }
    }
}