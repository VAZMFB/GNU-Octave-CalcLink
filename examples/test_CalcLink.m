% Test for CalcLink
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
close all, clear all, clc, tic
disp([' --- ' mfilename ' --- ']);

addpath([pwd '\..\']);

clApp = CalcLink(1); % Run LibreOffice Calc prgoram
% clApp.Open([pwd '\test.ods']); % Open existing file
clApp.New; % Create new file

% Add new sheet
clApp.AddSheet('Test');

% Remove sheet
clApp.RemoveSheet(2);

% Rename sheet
clApp.setSheetName(1, 'Test2');

% Write data
clApp.write(1, [1, 2, 1, 2], [2, 2, 3, 3], ...
  {'test', 'test2', '=5', '=2.454'});

% Read data
data = clApp.read(1, 2*ones(1, 5), 2:6);
disp(data);

% Change column widht and row height
clApp.setColumnWidth(1, 1, 50);
clApp.setRowHeight(1, 2, 30);

% Cell formating
clApp.format(1, 1, 2, 'CharWeight', 150, 'CharHeight', 16, ...
  'HoriJustify', 2)
clApp.format(1, 2, 2, 'Borders', [0.8, 255, 0, 0], ...
  'HoriJustify', 3, 'CharFontName', 'Arial', 'CharColor', [0, 255, 0])
clApp.format(1, [2, 3], [4, 4], 'CellBackColor', [255, 0, 0], ...
  'TopBorder', [0.4, 0, 0, 255])

clApp.SaveAs([pwd '\test.ods']); % Save document
% clApp.Quit(); % Close program

% - End of program
disp(' The program was successfully executed... ');
disp([' Execution time: ' num2str(toc, '%.2f') ' seconds']);
disp(' -------------------- ');