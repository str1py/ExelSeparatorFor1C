   M S B A M L   `   `   `  ����M  IExelSeparatorFor1C, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null7 clr-namespace:ExelSeparatorFor1CExelSeparatorFor1C  5   	   #    ExelSeparatorFor1C.MainWindow   5   	   \ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    5   	   R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5   	   @d2http://schemas.microsoft.com/expression/blend/2008     5   	   Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     5   	   2local clr-namespace:ExelSeparatorFor1C     5   	     � Title$  Excel��5   	   $	��350��6   $	��600��6$    � 
ResizeMode$ NoResize=�60      MainWindowStyle#���   6F   .��5	      �  � RowDefinitions 5
   
   �� 5      6   �� 5      6   �� 5      6   �� 5      6   �� 5      6   �� 5      6   �� 5      6   �� 5      6   �� 5      6   �� 5      6   �� 5      6    � ColumnDefinitions 5   
   �� 5      6!   �� 5      6!   �� 5      6!   �� 5      6!   �� 5      6!   �� 5      6!   .+�5   
   �� -   5         ��Name$ SelectButton��5      $��2��6'   $��5��64     MainButtonStyle#���  6`   .��6�   Выбрать файл6�   �� 5    
   -   5         ��Name$ 
FilePathTB��5       $��0��6&   $��2��66   $��5��6C     TextBoxStyle#���  6W   6�   �� 5!   
   -   5        $ DevideButton��5!      $��2��6'   $��2��67   $��1��6K   $��4��6\   #���  6�   .��6�   Разделить6�   �� 5"   
   -   5         ��Name$ 
LogListBox��5"      $��6��6&   $��4��63   $��1��6G   $��5��6W   
���������6h   Z� ��6�   Z� 6�   �� 5%   
   $��1q�6    �� CornerRadius$ 65��6&   $��Bottom=�68   $��6��6S   $��0��6h   $��0q�6u   $��2��6�     BorderStyle#���  6�   �� 5'   
   $��1q�6   $ 65��6&   $��6��69   $��5��6N   $��1��6[   $��Center=�6l   $��0q�6�   $��2��6�   #���  6�   5(      5)      