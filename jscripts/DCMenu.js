// --------------------------------------------------------------------------------
// Instantiate DCMenu
function DCMenu(MenuName, Top, Left, Width, UserRights, UserPermissions, Link, SelID)  {
    var SelIdx

    this.MenuName = MenuName;
    this.GetMenuName = function() {
        return this.MenuName
    }

    var MenuItems = [];
    this.MenuItems = MenuItems
    this.GetMenuItems = function() {
        return this.MenuItems 
    }
            
    this.GetTabNum = mitem_GetTabNum;
    function mitem_GetTabNum(id) {
        for (var i = 0; i <= MenuItems.length; i++) {
            if (MenuItems[i].id == id) {
                return i
            }
        }
    }
    
   // this.GetTabNum = function(id) {   
   //     return this.GetTabNum(id)
   // }
    
          
    
    // Assign methods
	//this.GetSomething = mitem_something;
	//function mitem_something(key) {
	//    return (key / 2)
	//}
    
    
    // ___ Build the MenuItems array //   
    switch(MenuName) {
        case "CallbackSetup":
            AppendMenuItems({id:"Home",       text:"Home",        RightRequired:'CBV',    PermissionRequired:"PUB"})        
            AppendMenuItems({id:"Setup",      text:"Setup",       RightRequired:'CBV',    PermissionRequired:"PUB"})
            AppendMenuItems({id:"Contact",    text:"Contact",     RightRequired:'CBV',    PermissionRequired:"OTHER"})
            //AppendMenuItems({id:"Authorize",  text:"Authorize",   RightRequired:'CBV',    PermissionRequired:'OTHER'})
            AppendMenuItems({id:"Notes",      text:"Notes",       RightRequired:'CBV',    PermissionRequired:'OTHER'})
            AppendMenuItems({id:"Escalate",   text:"Escalate",    RightRequired:'CBV',    PermissionRequired:'OTHER'}) 
            break;    
    }
        
    for (var i = 0; i < MenuItems.length; i++) {
        if (MenuItems[i].id == SelID) {
            SelIdx = i
            break;
        }
    }    

    // ___ Write the html //    
    document.write('<div id="' + MenuName + '" class = "DCMenu"> <ul class="DCMenuUL"> ')
    for (var i = 0 ; i < MenuItems.length ; i++) {
        if (i == SelIdx) {
            document.write('<li class = "selected"> ')
        } else {
            document.write('<li class = "unselected" > ')
        }   
        document.write('<b style="background-color: transparent;" class="rtop"> ')
        document.write('<b class="r1"></b><b class="r2"></b><b class="r3"></b><b class="r4"></b></b> ')
        document.write('<a href= "' + Link + '(' + i + ')">' + MenuItems[i].id + '</a> ')
        document.write('</li> ')
    }
    document.write('</ul></div>')  
    
    
    
   
          
    // ___ Set the top, left and width properties //
    var Name
    //var Form = document.getElementById("form1")
    var Form = document.getElementsByTagName("form")[0]
    var AllDiv = Form.getElementsByTagName("div")
    for ( i= 0 ; i < AllDiv.length ; i++ ) {
        Name = AllDiv[i].id
        if (Name == MenuName) {
        
            var Div = AllDiv[i]
            var DivWidth = ((parseFloat(Width) + 10) * MenuItems.length) + "px"; 
            
            Div.style.width = DivWidth
            //Div[i].style.margin = Top + " " + "0px 0px " + Left
            //Div.style.margin = Top + "px " + "0px 0px 0px"
            //Div.style.margin = "14px"
            Div.style.marginTop = parseInt(Top) + "px"
            Div.style.marginLeft = parseInt(Left) + "px"
            
            var AllUl = Div.getElementsByTagName("ul")
            var Ul = AllUl[0]
            
            //Ul.style.left="0px"
                    
            var AllLi = Ul.getElementsByTagName("li")
            for ( j = 0 ; j < AllLi.length ; j++ ) {
                var Li = AllLi[j]
                Li.style.width = Width
                
                Li.style.left= parseInt(j * 25) + "px"
                
                var AllA = Li.getElementsByTagName("a")
                A = AllA[0]
                A.style.width = Width                                           
            }                                  
        }
    }  
    
    // ___ AppendMenuItems //
    function AppendMenuItems(MenuItem) {
        var UserHasRequiredRight = 0
        var UserHasRequiredPermission = 0       
        
        // #1: Does user have the required right?
        var RightRequired_Array = MenuItem.RightRequired.toUpperCase().split("|")
        for (var i = 0; i < RightRequired_Array.length; i++) {   
            if (UserRights.toUpperCase().indexOf(RightRequired_Array[i]) >= 0) {
                UserHasRequiredRight = 1
                break;
            }   
        }    
      
        // #2: Does user have the required permission?
        if (UserHasRequiredRight == 1) {
            var PermissionRequired_Array = MenuItem.PermissionRequired.toUpperCase().split("|")
            for (var i = 0;i < PermissionRequired_Array.length;i++) {   
                if (UserPermissions.toUpperCase().indexOf(PermissionRequired_Array[i]) >= 0) {
                    UserHasRequiredPermission = 1
                    break;
                }   
            }         
        }
    
        // #3: Add menu items for which user has permissions to the MenuItemsRaw_Array
        if (UserHasRequiredRight == 1 && UserHasRequiredPermission == 1) {
                MenuItems[MenuItems.length] = {id: MenuItem.id, link: MenuItem.link}         
        }           
    }    
} 