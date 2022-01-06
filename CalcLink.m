% Class CalcLink
% Author: Milos D. Petrasinovic <mpetrasinovic@mas.bg.ac.rs>
% Structural Analysis of Flying Vehicles
% Faculty of Mechanical Engineering, University of Belgrade
% Department of Aerospace Engineering, Flying structures
% https://vazmfb.com
% Belgrade, 2021
% ---------------
%
% Copyright (C) 2021 Milos Petrasinovic <info@vazmfb.com>
%  
% This program is free software: you can redistribute it and/or modify
% it under the terms of the GNU General Public License as 
% published by the Free Software Foundation, either version 3 of the 
% License, or (at your option) any later version.
%   
% This program is distributed in the hope that it will be useful,
% but WITHOUT ANY WARRANTY; without even the implied warranty of
% MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
% GNU General Public License for more details.
%   
% You should have received a copy of the GNU General Public License
% along with this program.  If not, see <https://www.gnu.org/licenses/>.
%
% ---------------
classdef CalcLink < handle
  properties (Access = public)
    ServiceManager;
    Desktop;
    Document;
    Sheets;
  end
  properties (Access = private)
    flag;
    args;
    opened = 0;
    SheetNum = 0;
  end
  methods
    function self = CalcLink(mode)
      % mode - flag for program visibility
      
      if(nargin ~= 1 || (mode ~= 1 && mode ~= 0))
        mode = 1;
      end
      
      self.flag = (exist('OCTAVE_VERSION', 'builtin') > 0);
      if(self.flag)
        pkgFname = 'windows';
        min_version = '1.6.0';
        fpkg = pkg('list', pkgFname);
        if(~isempty(fpkg)) 
          if(nargin > 1)
            if(compare_versions(fpkg{1}.version, min_version, '>='))
              if(~fpkg{1}.loaded)
                pkg('load', pkgFname);
              end
            else
              disp(['Wait for ' pkgFname ' package to be updated...']);
              pkg('update', pkgFname);
              pkg('load', pkgFname);
              disp(' Package is updated and loaded...');
            end
          else
            if(~fpkg{1}.loaded)
              pkg('load', pkgFname);
            end
          end
        else
          disp(['Wait for ' pkgFname ' package to be installed...']);
          try
            pkg('install', '-forge', pkgFname);
            pkg('load', pkgFname);
            disp(' Package is installed and loaded...');
          catch
            error('Package installation failed!');
          end
        end

        warning('off', 'Octave:data-file-in-path');
      end
      try
        self.ServiceManager = actxserver('com.sun.star.ServiceManager');
        self.Desktop = invoke(self.ServiceManager, ...
          'createInstance', 'com.sun.star.frame.Desktop');

        self.args = invoke(self.ServiceManager, 'Bridge_GetStruct', ...
          "com.sun.star.beans.PropertyValue");
        name = invoke(self.ServiceManager, 'Bridge_GetValueObject');
        invoke(name, 'Set', 'string', 'Hidden');
        set(self.args, 'Name', name);
        val = invoke(self.ServiceManager, 'Bridge_GetValueObject');
        invoke(val, 'Set', 'boolean', ~mode);
        set(self.args, 'Value', val);
       catch
        error('Could not start LibreOffice!');
       end
    end
    
    function Visible(self, mode)
      % mode - flag for program visibility
      
      if(self.opened)
        if(mode == 1 || mode == 0)
          if(~isempty(self.Document))
            controller = get(self.Document, 'CurrentController');
            frame = get(controller, 'Frame');
            window = get(frame, 'ContainerWindow');
            invoke(window, 'setVisible', mode);
          end
          val = invoke(self.ServiceManager, 'Bridge_GetValueObject');
          invoke(val, 'Set', 'boolean', ~mode);
          set(self.args, 'Value', val);
        end
      else
        error('No opened document!');
      end
    end
    
    function Close(self)
      if(self.opened)
        invoke(self.Document, 'close', true);
        self.opened = 0;
      else
        error('No opened document!');
      end
    end
    
    function Quit(self)
      if(self.opened)
        invoke(self.Document, 'close', true);
      end
      invoke(self.Desktop, 'terminate');
      delete(self.Sheets);
      delete(self.Document);
      delete(self.Desktop);
      delete(self.ServiceManager);
      delete(self);
    end
    
    function Open(self, filePath)
      % filePath - full file path
      
      if(~self.opened)
        try
          self.opened = 1;
          urlString = self.pathToUrl(filePath);
          if(self.flag)
            windows_feature('COM_SafeArraySingleDim', 1);
            self.Document = invoke(self.Desktop, 'loadComponentFromURL', ...
              urlString, '_blank', 0, {self.args});
            windows_feature('COM_SafeArraySingleDim', 0);
          else
            feature('COM_SafeArraySingleDim', 1);
            self.Document = invoke(self.Desktop, 'loadComponentFromURL', ...
              urlString, '_blank', 0, {self.args});
            feature('COM_SafeArraySingleDim', 0);
          end 
          self.Sheets = invoke(self.Document, 'getSheets');
          self.SheetNum = invoke(self.Sheets, 'getCount');
       catch
        self.opened = 0;
        error('Could not open document!');
       end
      else
        disp('Document already opened!');
      end
    end

    function New(self)
      if(~self.opened)
        try
          self.opened = 1;
          if(self.flag)
            windows_feature('COM_SafeArraySingleDim', 1);
            self.Document = invoke(self.Desktop, 'loadComponentFromURL', ...
              'private:factory/scalc', '_blank', 0, {self.args});
            windows_feature('COM_SafeArraySingleDim', 0);
          else
            feature('COM_SafeArraySingleDim', 1);
            self.Document = invoke(self.Desktop, 'loadComponentFromURL', ...
              'private:factory/scalc', '_blank', 0, {self.args});
            feature('COM_SafeArraySingleDim', 0);
          end 
          self.Sheets = invoke(self.Document, 'getSheets');
          self.SheetNum = invoke(self.Sheets, 'getCount');
       catch
        self.opened = 0;
        error('Could not create new document!');
       end
      else
        disp('Document already opened!');
      end
    end
    
    function AddSheet(self, name)
      % name - new sheet name
      
      if(self.opened)
        if(ischar(name))
          if(~invoke(self.Sheets, 'hasByName', name))
              self.SheetNum = self.SheetNum+1;
              invoke(self.Sheets, 'insertNewByName', name, self.SheetNum-1);
          else
            error('Sheet with this name already exists!');
          end
        else
          error('Could not add sheet!');
        end
      else
        error('No opened document!');
      end
    end
    
    function RemoveSheet(self, sheet)
      % sheet - sheet index number
      
      if(self.opened)
        if(isnumeric(sheet) && sheet > 0 && sheet <= self.SheetNum)
          Sheet = invoke(self.Sheets, 'getByIndex', sheet-1);
          SheetName = get(Sheet, 'Name');
          invoke(self.Sheets, 'removeByName', SheetName);
          self.SheetNum = self.SheetNum-1;
        else
          error('Sheet does not exists!');
        end
      else
        error('No opened document!');
      end
    end
    
    function setSheetName(self, sheet, name)
      % sheet - sheet index number
      % name - new sheet name
      
      if(self.opened)
        if(isnumeric(sheet) && sheet > 0 && sheet <= self.SheetNum)
          if(ischar(name))
            if(~invoke(self.Sheets, 'hasByName', name))
              Sheet = invoke(self.Sheets, 'getByIndex', sheet-1);
              set(Sheet, 'Name', name);
            else
              error('Sheet with this name already exists!');
            end
          else
            error('Could not add sheet!');
          end
        else
          error('Sheet does not exist!');
        end
      else
        error('No opened document!');
      end
    end
    
    function write(self, sheet, col, row, data)
      % sheet - sheet index number
      % col - cell column index number
      % row - cell row index number
      % data - data for write operation
      
      if(self.opened)
        if(isnumeric(sheet) && sheet > 0 && sheet <= self.SheetNum)
          if(length(col) == length(row))
            if(~iscell(data))
                data = {data};
            end
            for i = 1:length(row)
              if(isnumeric(col) && isnumeric(row) && ...
                  all(col > 0) && all(row > 0))
                Sheet = invoke(self.Sheets, 'getByIndex', sheet-1);
                Cell = invoke(Sheet, 'getCellByPosition', col(i)-1, ...
                  row(i)-1);
                invoke(Cell, 'setFormula', data{i});
              else
                error('Invalid cell position!');
              end
            end
          else
            error('Different dimensions of column and row vectors!');
          end
        else
          error('Sheet does not exist!');
        end
      else
        error('No opened document!');
      end
    end
    
    function data = read(self, sheet, col, row)
      % sheet - sheet index number
      % col - cell column
      % row - cell row
      % data - data after read operation
      
      data = {};
      if(self.opened)
        if(isnumeric(sheet) && sheet > 0 && sheet <= self.SheetNum)
          if(length(col) == length(row))
            for i = 1:length(row)
              if(isnumeric(col) && isnumeric(row) && ...
                  all(col > 0) && all(row > 0))
                Sheet = invoke(self.Sheets, 'getByIndex', sheet-1);
                Cell = invoke(Sheet, 'getCellByPosition', col(i)-1, ...
                  row(i)-1);
                data{i}.Value = invoke(Cell, 'getValue');
                data{i}.String = invoke(Cell, 'getString');
                data{i}.Formula = invoke(Cell, 'getFormula');
              else
                error('Invalid cell position!');
              end
            end
          else
            error('Different dimensions of column and row vectors!');
          end
        else
          error('Sheet does not exist!');
        end
      else
        error('No opened document!');
      end
    end
    
    function setColumnWidth(self, sheet, col, width)
      % sheet - sheet index number
      % col - column index number
      % width - coulumn width
      
      if(self.opened)
        if(isnumeric(sheet) && sheet > 0 && sheet <= self.SheetNum)
          if(isnumeric(col) && isnumeric(width) &&...
               all(col > 0) && all(width > 0) &&...
               length(col) == length(width))
            Sheet = invoke(self.Sheets, 'getByIndex', sheet-1);
            Columns = invoke(Sheet, 'getColumns');
            for i = 1:length(col)
              oCol = invoke(Columns, 'getByIndex', col(i)-1);
              set(oCol, 'Width', width(i)*100);
            end
          else
            error('Could not set row height!');
          end
        else
          error('Sheet does not exist!');
        end
      else
        error('No opened document!');
      end
    end
    
    function setRowHeight(self, sheet, row, height)
      % sheet - sheet index number
      % row - row index number
      % height - row height
      
      if(self.opened)
        if(isnumeric(sheet) && sheet > 0 && sheet <= self.SheetNum)
          if(isnumeric(row) && isnumeric(height) &&...
               all(row > 0) && all(height > 0) &&...
               length(row) == length(height))
            Sheet = invoke(self.Sheets, 'getByIndex', sheet-1);
            Rows = invoke(Sheet, 'getRows');
            for i = 1:length(row)
              oRow = invoke(Rows, 'getByIndex', row(i)-1);
              set(oRow, 'Height', height(i)*100);
            end
          else
            error('Could not set row height!');
          end
        else
          error('Sheet does not exist!');
        end
      else
        error('No opened document!');
      end
    end
    
    function format(self, sheet, col, row, varargin)
      % sheet - sheet index number
      % col - cell column index number
      % row - cell row index number
      % varargin - additional variables
      
      % Documentation: 
      % https://www.openoffice.org/
      % api/docs/common/ref/com/sun/star/table/CellProperties.html
      % https://www.openoffice.org/
      % api/docs/common/ref/com/sun/star/style/CharacterProperties.html
      
      % http://www.openoffice.org/
      % api/docs/common/ref/com/sun/star/awt/FontWeight.html
      % http://www.openoffice.org/
      % api/docs/common/ref/com/sun/star/table/CellVertJustify.html
      % http://www.openoffice.org/
      % api/docs/common/ref/com/sun/star/table/CellHoriJustify.html
      % http://www.openoffice.org/
      % api/docs/common/ref/com/sun/star/table/CellOrientation.html
      % http://www.openoffice.org/
      % api/docs/common/ref/com/sun/star/awt/FontSlant.html
      
      if(self.opened)
        if(isnumeric(sheet) && sheet > 0 && sheet <= self.SheetNum)
          if(length(col) == length(row))
            rgb = @(r, g, b) b*1+g*256+r*256^2;
            p = inputParser;
            addParameter(p, 'CellBackColor', [], ...
              @(x) isnumeric(x) && length(x)==3);
            addParameter(p, 'CharWeight', [], ...
              @(x) isnumeric(x) && length(x)==1);
            addParameter(p, 'CharPosture', [], ...
              @(x) isnumeric(x) && length(x)==1);
            addParameter(p, 'CharFontName', [], ...
              @(x) ischar(x));
            addParameter(p, 'CharHeight', [], ...
              @(x) isnumeric(x) && length(x)==1);
            addParameter(p, 'CharColor', [], ...
              @(x) isnumeric(x) && length(x)==3);
            addParameter(p, 'HoriJustify', [], ...
              @(x) isnumeric(x) && length(x)==1);
            addParameter(p, 'VertJustify', [], ...
              @(x) isnumeric(x) && length(x)==1);
            addParameter(p, 'RotateAngle', [], ...
              @(x) isnumeric(x) && length(x)==1);
            addParameter(p, 'Orientation', [], ...
              @(x) isnumeric(x) && length(x)==1);
            addParameter(p, 'Borders', [], ...
              @(x) isnumeric(x) && length(x)==4);
            addParameter(p, 'TopBorder', [], ...
              @(x) isnumeric(x) && length(x)==4);
            addParameter(p, 'BottomBorder', [], ...
              @(x) isnumeric(x) && length(x)==4);
            addParameter(p, 'LeftBorder', [], ...
              @(x) isnumeric(x) && length(x)==4);
            addParameter(p, 'RightBorder', [], ...
              @(x) isnumeric(x) && length(x)==4);
            p.parse(varargin{:})
          
            for i = 1:length(row)
              if(isnumeric(col) && isnumeric(row) && ...
                  all(col > 0) && all(row > 0))
                Sheet = invoke(self.Sheets, 'getByIndex', sheet-1);
                Cell = invoke(Sheet, 'getCellByPosition', col(i)-1, ...
                  row(i)-1);
                
                if(~isempty(p.Results.CellBackColor)) 
                  c = p.Results.CellBackColor;
                  set(Cell, 'CellBackColor', rgb(c(1), c(2), c(3)));
                end
                if(~isempty(p.Results.CharWeight)) 
                  set(Cell, 'CharWeight', p.Results.CharWeight);
                end
                if(~isempty(p.Results.CharPosture)) 
                  set(Cell, 'CharPosture', p.Results.CharPosture);
                end
                if(~isempty(p.Results.CharFontName)) 
                  set(Cell, 'CharFontName', p.Results.CharFontName);
                end
                if(~isempty(p.Results.CharHeight)) 
                  set(Cell, 'CharHeight', p.Results.CharHeight);
                end
                if(~isempty(p.Results.CharColor)) 
                  c = p.Results.CharColor;
                  set(Cell, 'CharColor', rgb(c(1), c(2), c(3)));
                end
                if(~isempty(p.Results.HoriJustify)) 
                  set(Cell, 'HoriJustify', p.Results.HoriJustify);
                end
                if(~isempty(p.Results.VertJustify)) 
                  set(Cell, 'VertJustify', p.Results.VertJustify);
                end
                if(~isempty(p.Results.RotateAngle)) 
                  set(Cell, 'RotateAngle', p.Results.RotateAngle);
                end
                if(~isempty(p.Results.Orientation)) 
                  set(Cell, 'Orientation', p.Results.Orientation);
                end
                if(~isempty(p.Results.Borders))
                  w = p.Results.Borders(1)*100; 
                  c = p.Results.Borders(2:4); 
                  crgb = rgb(c(1), c(2), c(3));
                  Border = get(Cell, 'TopBorder');
                  set(Border, 'Color', crgb);
                  set(Border, 'InnerLineWidth', w);
                  set(Cell, 'TopBorder', Border);
                  set(Cell, 'BottomBorder', Border);
                  set(Cell, 'LeftBorder', Border);
                  set(Cell, 'RightBorder', Border);
                end
                if(~isempty(p.Results.TopBorder)) 
                  w = p.Results.TopBorder(1)*100; 
                  c = p.Results.TopBorder(2:4); 
                  Border = get(Cell, 'TopBorder');
                  set(Border, 'Color', rgb(c(1), c(2), c(3)));
                  set(Border, 'InnerLineWidth', w);
                  set(Cell, 'TopBorder', Border);
                end
                if(~isempty(p.Results.BottomBorder)) 
                  w = p.Results.BottomBorder(1)*100; 
                  c = p.Results.BottomBorder(2:4); 
                  Border = get(Cell, 'TopBorder');
                  set(Border, 'Color', rgb(c(1), c(2), c(3)));
                  set(Border, 'InnerLineWidth', w);
                  set(Cell, 'BottomBorder', Border);
                end
                if(~isempty(p.Results.LeftBorder)) 
                  w = p.Results.LeftBorder(1)*100; 
                  c = p.Results.LeftBorder(2:4); 
                  Border = get(Cell, 'TopBorder');
                  set(Border, 'Color', rgb(c(1), c(2), c(3)));
                  set(Border, 'InnerLineWidth', w);
                   set(Cell, 'LeftBorder', Border);
                end
                if(~isempty(p.Results.RightBorder)) 
                  w = p.Results.RightBorder(1)*100; 
                  c = p.Results.RightBorder(2:4); 
                  Border = get(Cell, 'TopBorder');
                  set(Border, 'Color', rgb(c(1), c(2), c(3)));
                  set(Border, 'InnerLineWidth', w);
                  set(Cell, 'RightBorder', Border);
                end
              else
                error('Invalid cell position!');
              end
            end
          else
            error('Different dimensions of column and row vectors!');
          end
        else
          error('Sheet does not exist!');
        end
      else
        error('No opened document!');
      end
    end
    
    function Save(self)
      if(self.opened)
        invoke(self.Document, 'store');
      else
        error('No opened document!');
      end
    end

    function SaveAs(self, filePath)
      % filePath - full file path
      
      if(self.opened)
        urlString = self.pathToUrl(filePath);
        
        prop = invoke(self.ServiceManager, 'Bridge_GetStruct', ...
          "com.sun.star.beans.PropertyValue");
        name = invoke(self.ServiceManager, 'Bridge_GetValueObject');
        invoke(name, 'Set', 'string', 'FilterName');
        set(prop, 'Name', name);
        val = invoke(self.ServiceManager, 'Bridge_GetValueObject');
        invoke(val, 'Set', 'string', 'calc8');
        set(prop, 'Value', val);
        
        if(exist(filePath, 'file') == 2)
          delete(filePath);
        end
        
        if(self.flag)
          windows_feature('COM_SafeArraySingleDim', 1);
          invoke(self.Document, 'storeAsURL', urlString, {prop});
          windows_feature('COM_SafeArraySingleDim', 0);
        else
          feature('COM_SafeArraySingleDim', 1);
          invoke(self.Document, 'storeAsURL', urlString, {prop});
          feature('COM_SafeArraySingleDim', 0);
        end 
      else
        error('No opened document!');
      end
    end
    
    function urlString = pathToUrl(~, filePath)
      % filePath - full file path
      % urlString - file url
      
      % Convert filepath to url
      char_in = {'%', '\', '#', '{', '}', '`', '^', ';', '@', '[', ']', ...
        ''''};
      char_out = {'%25', '/', '%23', '%7B', '%7D', '%C2%B4', '%5E', ...
        '%3b', '%40', '%5B', '%5D', '%27'};
      for i = 1:length(char_in)
        filePath = strrep(filePath, char_in{i}, char_out{i});
      end
      urlString = ['file:///' filePath];
    end
  end
end
    