   M S B A M L   `   `   `  ����  ExcelComparator     ExcelComparator.MainWindow   5   	   W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    5   	   R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5   	   @d2http://schemas.microsoft.com/expression/blend/2008     5   	   Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     5   	   U QMaterialDesignThemes.Wpf, Version=2.3.1.953, Culture=neutral, PublicKeyToken=nullFmaterialDesign1http://materialdesigninxaml.net/winfx/xaml/themes  5   	      MaterialDesignBody#���   5   	   $��Regular"�5   	   $��15(�5	   	   &   System.Windows.Media.TextOptions    TextFormattingMode$  Display=�5
   	      TextRenderingMode$
 Auto=�5   	     MaterialDesignPaper#���  5   	     MaterialDesignFont#���  5   	   $
��Auto��5   	   $
��Auto��6    � SizeToContent$ WidthAndHeight=�6$    � 
ResizeMode$ NoResize=�5   	    � WindowStartupLocation$ CenterScreen=�5   	    � Title$ Excel Comparator��5   	   .��5      ��  �� Orientation$ Vertical=�6   $��5q�6(   .+�5   
   #  MaterialDesignThemes.Wpf.Card  $��10q�6   $��5q�6+   .��5      �� $��10q�6   $ Vertical=�6&   $��Center=�6=   .+�5      �� .��6   Path File 16%   �� 5      -   5        $��file1��5   '   $
��auto��6   66   �� 5      .��6   Sheet Name (default first)64   �� 5      -   5        $��sheet��5      6)   �� 5      -   5        $��5q�5      .��69   Browse6A   5      5        5   
   $��10q�6   $��5q�6+   .��5      �� $��10q�6   $
��auto��6%   $ Vertical=�62   $��Center=�6I   .+�5      �� .��6   Path File 26%   �� 5       -   5        $��file2��5       6)   �� 5!      .��6   Sheet Name (default first)64   �� 5"      -   5        $��sheet2��5"      6*   �� 5#      -   5        $��5q�5#      .��6;   Browse6C   5$      5%        5&   
   $��10q�6   $��5q�6+   .��5'      �� $��10q�6   $
��auto��6%   $ Vertical=�62   $��Center=�6I   .+�5(      �� .��6   31Column To Compare (separe with ';') (default all)6K   �� 5)      -   5        $��columns��5)      6+   5*      5+      �� 5,   
   -   5        $��10q�5,      $
��auto��6   .��6A   
Validate6K   5-      5.      