using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TimeTracker.Entity
{
    public class Session
    {
        /// <summary>
        /// The x coordinate for adding controls
        /// </summary>
        public int x { get; set; }  

        /// <summary>
        /// The y coorodinate for adding controls
        /// </summary>
        public int y { get; set; } 

        /// <summary>
        /// The index is used for iterating through the controls
        /// </summary>
        public int index { get; set; } 

        /// <summary>
        /// The index of the active line
        /// </summary>
        public int activeIndex { get; set; } 

        /// <summary>
        /// Keeps track of the total amount of lines
        /// </summary>
        public int totalLines { get; set; } 

        /// <summary>
        /// The state of the timer
        /// </summary>
        public bool isTimerActive { get; set; } 

        /// <summary>
        /// This hold the current file directory
        /// </summary>
        public string fileDirectory { get; set; } 
    }
}
