using System.Collections;
using System.Collections.Generic;

namespace GameInfos.Models
{
    public class GameModel
    {
        public GameModel()
        {
            InstallCount = new List<string>();
            GameName = new List<string>();
            CompanyName = new List<string>();
            Email = new List<string>();
            Country = new List<string>();
            
        }
        
      
        

        public List<string> InstallCount { get; set; }
        public List<string> GameName { get; set; }
        public List<string> CompanyName { get; set; }
        public List<string> Email { get; set; }
        public List<string> Country { get; set; }
       
    }
}