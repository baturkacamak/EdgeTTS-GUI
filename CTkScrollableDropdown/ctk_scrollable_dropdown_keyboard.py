'''
Enhanced Scrollable Dropdown class with keyboard navigation support
'''

from .ctk_scrollable_dropdown import CTkScrollableDropdown
import customtkinter
import difflib

class CTkScrollableDropdownKeyboard(CTkScrollableDropdown):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        
        # Keep track of currently selected index
        self.selected_index = -1
        self.visible_widgets = []  # List to keep track of currently visible widgets
        
        # Bind keyboard events to the combobox entry
        if isinstance(self.attach, customtkinter.CTkComboBox):
            self.attach._entry.bind("<Up>", self._on_up_key, add="+")
            self.attach._entry.bind("<Down>", self._on_down_key, add="+")
            self.attach._entry.bind("<Return>", self._on_return_key, add="+")
            
    def _on_up_key(self, event):
        """Handle Up arrow key press"""
        if not self.winfo_viewable() or not self.visible_widgets:
            return
        
        # Update selected index
        if self.selected_index == -1:
            self.selected_index = len(self.visible_widgets) - 1
        else:
            self.selected_index = (self.selected_index - 1) % len(self.visible_widgets)
            
        self._update_selection()
        return "break"  # Prevent default handling
        
    def _on_down_key(self, event):
        """Handle Down arrow key press"""
        if not self.winfo_viewable() or not self.visible_widgets:
            return
            
        # Update selected index
        if self.selected_index == -1:
            self.selected_index = 0
        else:
            self.selected_index = (self.selected_index + 1) % len(self.visible_widgets)
            
        self._update_selection()
        return "break"  # Prevent default handling
        
    def _on_return_key(self, event):
        """Handle Return key press"""
        if not self.winfo_viewable() or not self.visible_widgets:
            return
            
        if 0 <= self.selected_index < len(self.visible_widgets):
            # Get the text of the selected widget and trigger selection
            selected_text = self.visible_widgets[self.selected_index].cget("text")
            self._attach_key_press(selected_text)
        return "break"  # Prevent default handling
        
    def _update_selection(self):
        """Update the visual selection in the dropdown"""
        # Reset all buttons to default color
        for widget in self.visible_widgets:
            widget.configure(fg_color=self.button_color)
            
        # Highlight selected button
        if 0 <= self.selected_index < len(self.visible_widgets):
            self.visible_widgets[self.selected_index].configure(fg_color=self.hover_color)
            
    def live_update(self, string=None):
        """Override live_update to keep track of visible widgets"""
        if not self.appear or self.disable or self.fade:
            return
            
        self.selected_index = -1  # Reset selection on new search
        self.visible_widgets = []  # Clear visible widgets list
        
        if string:
            string = string.lower()
            self._deiconify()
            i = 1
            for key in self.widgets.keys():
                s = self.widgets[key].cget("text").lower()
                text_similarity = difflib.SequenceMatcher(None, s[0:len(string)], string).ratio()
                similar = s.startswith(string) or text_similarity > 0.75
                if not similar:
                    self.widgets[key].pack_forget()
                else:
                    self.widgets[key].pack(fill="x", pady=2, padx=(self.padding, 0))
                    self.visible_widgets.append(self.widgets[key])
                    i += 1
                    
            if i == 1:
                self.no_match.pack(fill="x", pady=2, padx=(self.padding, 0))
            else:
                self.no_match.pack_forget()
            self.button_num = i
            self.place_dropdown()
        else:
            self.no_match.pack_forget()
            self.button_num = len(self.values)
            for key in self.widgets.keys():
                self.widgets[key].destroy()
            self._init_buttons()
            self.visible_widgets = list(self.widgets.values())
            self.place_dropdown()
            
        self.frame._parent_canvas.yview_moveto(0.0)
        self.appear = False 