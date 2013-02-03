function varargout = PopupCamera(varargin)
% POPUPCAMERA MATLAB code for PopupCamera.fig
%      POPUPCAMERA, by itself, creates a new POPUPCAMERA or raises the existing
%      singleton*.
%
%      H = POPUPCAMERA returns the handle to a new POPUPCAMERA or the handle to
%      the existing singleton*.
%
%      POPUPCAMERA('CALLBACK',hObject,eventData,handles,...) calls the local
%      function named CALLBACK in POPUPCAMERA.M with the given input arguments.
%
%      POPUPCAMERA('Property','Value',...) creates a new POPUPCAMERA or raises the
%      existing singleton*.  Starting from the left, property value pairs are
%      applied to the GUI before PopupCamera_OpeningFcn gets called.  An
%      unrecognized property name or invalid value makes property application
%      stop.  All inputs are passed to PopupCamera_OpeningFcn via varargin.
%
%      *See GUI Options on GUIDE's Tools menu.  Choose "GUI allows only one
%      instance to run (singleton)".
%
% See also: GUIDE, GUIDATA, GUIHANDLES

% Edit the above text to modify the response to help PopupCamera

% Last Modified by GUIDE v2.5 28-Jan-2013 18:27:33

% Begin initialization code - DO NOT EDIT
gui_Singleton = 1;
gui_State = struct('gui_Name',       mfilename, ...
                   'gui_Singleton',  gui_Singleton, ...
                   'gui_OpeningFcn', @PopupCamera_OpeningFcn, ...
                   'gui_OutputFcn',  @PopupCamera_OutputFcn, ...
                   'gui_LayoutFcn',  [] , ...
                   'gui_Callback',   []);
if nargin && ischar(varargin{1})
    gui_State.gui_Callback = str2func(varargin{1});
end

if nargout
    [varargout{1:nargout}] = gui_mainfcn(gui_State, varargin{:});
else
    gui_mainfcn(gui_State, varargin{:});
end
% End initialization code - DO NOT EDIT


% --- Executes just before PopupCamera is made visible.
function PopupCamera_OpeningFcn(hObject, eventdata, handles, varargin)
% This function has no output args, see OutputFcn.
% hObject    handle to figure
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
% varargin   command line arguments to PopupCamera (see VARARGIN)

% Choose default command line output for PopupCamera
handles.output = hObject;

% Update handles structure
guidata(hObject, handles);





% --- Outputs from this function are returned to the command line.
function varargout = PopupCamera_OutputFcn(hObject, eventdata, handles) 
% varargout  cell array for returning output args (see VARARGOUT);
% hObject    handle to figure
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Get default command line output from handles structure
varargout{1} = handles.output;


% --- Executes on selection change in Camera.
function Camera_Callback(hObject, eventdata, handles)
set(handles.uitable1,'data',[]);
axes(handles.axes1);
check=get(handles.Camera,'value');
if (check==1)
    obj=videoinput('winvideo',1);
else(check==2)
    obj=videoinput('winvideo',2,'YUY2_640x480');
end
vidRes = get(obj, 'VideoResolution'); % Resolution of device ex. 640x480
nBands = get(obj, 'NumberOfBands'); % number of channels ex. 3=RGB image
hImage = image( zeros(vidRes(2), vidRes(1), nBands) ); 
preview(obj, hImage);
handles.objVideo=obj;
guidata(hObject, handles);

% --- Executes during object creation, after setting all properties.
function Camera_CreateFcn(hObject, eventdata, handles)
% hObject    handle to Camera (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: popupmenu controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end


% --- Executes on button press in Connect.
function Connect_Callback(hObject, eventdata, handles)

checks = get(handles.chack,'value');
if (checks==1)
    
global couseAns ;
couseAns = zeros(1,4);
global scorAns ;
scorAns = zeros(30,4);
    
im=getsnapshot(handles.objVideo);
I=im;
%imshow(I);
Igray = rgb2gray(I);
Iinvert = 255-Igray;
J = imadjust(Iinvert);
H = fspecial('average',25);
I = imfilter(I,H,'replicate');
level = graythresh(J);
Ibw = im2bw(J,level);
Iopen = bwareaopen(Ibw,5);
[L n] = bwlabel(Iopen);
S = regionprops(L,'all');
box1 = S(5).BoundingBox ;
box2 = S(15).BoundingBox ;
box3 = S(17).BoundingBox ;
box4 = S(47).BoundingBox ;
R1 = box1(1,2);
R2 = box2(1,2);
R3 = box3(1,2);
R4 = box4(1,2);
IABWcrop2 = Iopen(R1:R2,128:175); % รหัสวิชา
IABWcrop3 = Iopen(R3:R4,35:85);

%--------------------------------รหัสวิชา----------------------------------
scor2  = zeros(10,4);
answer2 = zeros(10,4);
i3=0;
j3=0;
for i=0:9   % วนรหัสวิชา
    i3=i3+1;    
      for j = 0:3
        j3=j3+1;  
        W2 = IABWcrop2(1+(7.5*i):7.5*(i+1), 11.3*j+1 : 11.3*j + 11.3);           
        
        count2 = sum(sum(W2));       
        if count2> 5
            answer2(i3,j3) = 1;             
        else
            answer2(i3,j3) = 0;           
        end
         scor2(i3,j3) = answer2(i3,j3);        
      end
      j3=0;
end
num2 = zeros(10,4);
for i=1:10    % วนกำหนดตัวเลขตามข้อ
   for j=1:4
      if scor2(i,j)==1 
          num2(i,j) = i;
         if i == 10
            if scor2(i,j)==1
             num2(i,j)= 0; 
            end
         end
      end
   end
end

%couseAns = zeros(1,4);
for i=1:10  % วนเอาเลขมาเรียงต่อกัน       
   for j=1:4
       if num2(i,j)>=1
          couseAns(j) = num2(i,j); 
       end
   end
end
%msgbox(num2str(couseAns));
set(handles.show,'string',num2str(couseAns));
%-----------------------------ช่วงตรวจข้อสอบ-------------------------------
 
%scorAns  = zeros(30,4);
answer3 = zeros(30,4);

i4=0;
j4=0;
for i=0:29
    i4=i4+1;    
      for j = 0:3
        j4=j4+1;  
       W4 = IABWcrop3(1+(8.1*i):8.1*(i+1), 11.2*j+1 : 11.2*j + 11.2);
        count4 = sum(sum(W4));       
        if count4> 10
            answer3(i4,j4) = 1;             
        else
            answer3(i4,j4) = 0;           
        end
         scorAns(i4,j4) = answer3(i4,j4);        
      end
      j4=0;
end
%msgbox(num2str(scorAns));
msgbox('บันทึกแบบเฉลยแล้ว','บันทึก');
end
%  mastercouse = couseAns;
%  masterscore = scorAns;
% msgbox(num2str(couseAns));
%-----------------------------ตรวจแบบทดสอบ-------------------------------
if(checks==2)   
    
global couseAns ;
global scorAns ;
    
   %msgbox(num2str(couseAns));  
  %msgbox(num2str(scorAns));
im=getsnapshot(handles.objVideo);
I=im;
%imshow(I);
Igray = rgb2gray(I);
Iinvert = 255-Igray;
J = imadjust(Iinvert);
H = fspecial('average',25);
I = imfilter(I,H,'replicate');
level = graythresh(J);
Ibw = im2bw(J,level);
Iopen = bwareaopen(Ibw,5);
[L n] = bwlabel(Iopen);
S = regionprops(L,'all');
box1 = S(5).BoundingBox ;
box2 = S(15).BoundingBox ;
box3 = S(17).BoundingBox ;
box4 = S(47).BoundingBox ;

C1 = box1(1,1); % ก้อนที่ 5
R1 = box1(1,2);

C2 = box2(1,1); % ก้อนที่ 15
R2 = box2(1,2);

C3 = box3(1,1); % ก้อนที่ 17
R3 = box3(1,2);

C4 = box4(1,1); % ก้อนที่ 47
R4 = box4(1,2);

IABWcrop = Iopen(R1:R2,35:115); % รหัสประจำตัว
IABWcrop2 = Iopen(R1:R2,128:172); % รหัสวิชา
IABWcrop3 = Iopen(R3:R4,35:85);
%------------------------------รหัสประจำตัว--------------------------------
scor  = zeros(10,7);
answer = zeros(10,7);
i2=0;
j2=0; % วนรหัสประจำตัว
for i=0:9
    i2=i2+1;    
      for j = 0:6
        j2=j2+1;  
        W = IABWcrop(1+(7.4*i):7.4*(i+1), 11.5*j+1 : 11.5*j + 11.5); 
          
        count1 = sum(sum(W));       
        if count1> 10
            answer(i2,j2) = 1;             
        else
            answer(i2,j2) = 0;           
        end
         scor(i2,j2) = answer(i2,j2);       
      end
      j2=0;
end

num = zeros(10,7);
for i=1:10     % กำหนดเลขตามข้อ
   for j=1:7
      if scor(i,j)==1 
          num(i,j) = i;
         if i == 10
            if scor(i,j)==1
             num(i,j)= 0; 
            end
         end
      end
   end
end

number = zeros(1,7);
for i=1:10  % เอาเลขมาต่อเรียงกันเป็นรหัสประจำตัว
   for j=1:7
       if num(i,j)>=1
          number(j) = num(i,j); 
       end
   end
end
%msgbox(num2str(number));
%---------------------------------รหัสวิชา--------------------------------
scor2  = zeros(10,4);
answer2 = zeros(10,4);

i3=0;
j3=0;
for i=0:9   % วนรหัสวิชา
    i3=i3+1;    
      for j = 0:3
        j3=j3+1;  
        W2 = IABWcrop2(1+(7.5*i):7.5*(i+1), 11.3*j+1 : 11.3*j + 11.3);           
        
        count2 = sum(sum(W2));       
        if count2> 5
            answer2(i3,j3) = 1;             
        else
            answer2(i3,j3) = 0;           
        end
         scor2(i3,j3) = answer2(i3,j3);        
      end
      j3=0;
end
num2 = zeros(10,4);
for i=1:10    % วนกำหนดตัวเลขตามข้อ
   for j=1:4
      if scor2(i,j)==1 
          num2(i,j) = i;
         if i == 10
            if scor2(i,j)==1
             num2(i,j)= 0; 
            end
         end
      end
   end
end

couseTest = zeros(1,4);
for i=1:10  % วนเอาเลขมาเรียงต่อกัน       
   for j=1:4
       if num2(i,j)>=1
          couseTest(j) = num2(i,j); 
       end
   end
end
%msgbox(num2str(couseTest));
%----------------------------เปรียบเทียบรหัสวิชา------------------------------
idcouse = 0;
for i=1:1
      for j = 1:4 
           compar = couseTest(i,j) == couseAns(i,j);
        if compar
          idcouse = idcouse+1;          
%           if idcouse == 4      
%              msgbox('รหัสวิชาถูกต้อง');          
%           end         
        end 
        
      end 
      if idcouse < 4 
             msgbox('รหัสวิชาไม่ถูกต้อง'); 
             pause
      end
    idcouse=0;
end 
    
%------------------------------ช่วงตรวจข้อสอบ--------------------------------
 
scorTests  = zeros(30,4);
answer3 = zeros(30,4);

i4=0;
j4=0;
for i=0:29
    i4=i4+1;
    
      for j = 0:3
        j4=j4+1;  
        W4 = IABWcrop3(1+(8.1*i):8.1*(i+1), 11.2*j+1 : 11.2*j + 11.2);
        count4 = sum(sum(W4));       
        if count4> 10
            answer3(i4,j4) = 1;             
        else
            answer3(i4,j4) = 0;           
        end
         scorTests(i4,j4) = answer3(i4,j4);        
      end
      j4=0;
end


%-----------------------------เปรียบเทียบข้อสอบ------------------------------
%msgbox(num2str(scorTests));
point = 0;
score = 0;
for i=1:30
      for j = 1:4 
           compar = scorTests(i,j) == scorAns(i,j);
        if compar
          point = point+1;
          if point == 4              
              score = score+1;
          end
        end  
      end     
    point=0;
end


identity = num2str(number);
identity(identity ==' ') = '';
niden = str2num(identity);

course = num2str(couseTest);
course(course ==' ') = '';
ncourse = str2num(course);

checkAns = num2str(score);
checkAns(checkAns ==' ')='';
nscore = str2num(checkAns);

%people = 0;
totle = zeros(1,3);
if totle(1,1) == 0;
    totle(1,1) = totle(1,1) + niden;
    %msgbox(num2str(totle(1,1)));
end
if totle(1,2) == 0;           
    totle(1,2) = totle(1,2) + ncourse;
end
if totle(1,3) == 0;
   totle(1,3) = totle(1,3) + nscore; 
end 
 A = num2str(get(handles.uitable1,'data')); 
 B = num2str(totle);
 C = strvcat(A,B);
 
 % msgbox(num2str(size(A,1)));
   %msgbox(num2str(A(1:7)));
   %msgbox(num2str(B));
   %msgbox(num2str(C));
%    for i=1:size(A,1)
%        AA = A(i,:);
%        AAA = AA(i:7);
%    end
   if A ~= 0       
       for i=1:size(A,1)
           TTT = A(i,:);
           TTTT = TTT(1:7);
           if strcmp(B(1:7),TTTT)  
               
                button = questdlg('รหัสนักศึกษาซ้ำ คุณต้องการเปลี่ยนแปลงหรือไม่?','รหัสนักศึกษาซ้ำ','Yes','No','No');
                switch button
                    case 'Yes',       
                       D = unique(C,'rows');
                        break
                        
                    case 'No',
                        pause
                end
           end
            
       end
   end

D = unique(C,'rows');
sort = sortrows(D);
%msgbox(num2str(sort));
set(handles.uitable1,'data',str2num(sort)); 
msgbox('เสร็จสิ้น','เสร็จสิ้น');
end
%end

guidata(hObject,handles);


% --- Executes on button press in Processing.
function Processing_Callback(hObject, eventdata, handles)
button = questdlg('Do you want to quit?', ...
'Exit Dialog','Yes','No','No');
switch button
case 'Yes',
disp('Exit RGB SpecAnal');
%Save variables to matlab.mat
save
close(ancestor(hObject,'figure'))
case 'No',
quit cancel;
end

% --- Executes on selection change in chack.
function chack_Callback(hObject, eventdata, handles)

function chack_CreateFcn(hObject, eventdata, handles)


% Hint: popupmenu controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end
function show_Callback(hObject, eventdata, handles)
% --- Executes during object creation, after setting all properties.
function show_CreateFcn(hObject, eventdata, handles)


% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end


% --- Executes on button press in reset.
function reset_Callback(hObject, eventdata, handles)
set(handles.uitable1,'data',[]);


% --- Executes on button press in pushbutton5.
function pushbutton5_Callback(hObject, eventdata, handles)
closepreview


% --- Executes when entered data in editable cell(s) in uitable1.
function uitable1_CellEditCallback(hObject, eventdata, handles)


% --- Executes on button press in explot.
function explot_Callback(hObject, eventdata, handles)
SelecY = get(handles.uitable1,'data');
nrow = size(SelecY,1);
C = cell(3,3);
C{1,1} = 'รหัสนักศึกษา';
C{1,2} = 'รหัสวิชา';
C{1,3} = 'คะแนน';

for i=1:nrow
    C{i+1,1} = SelecY(i,1);
    C{i+1,2} = SelecY(i,2);
    C{i+1,3} = SelecY(i,3);
end

FileName = uiputfile('*.xls','Save as');
        xlswrite(FileName,C),
