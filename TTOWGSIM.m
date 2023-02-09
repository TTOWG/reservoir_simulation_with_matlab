function ResPerfPara = TTOWGSIM(DataDeck)
    %TTOWG - To The Only Wise God be Glory!!!
    %ResPerfPara = TTOWGSIM(DataDeck) returns the 2-D simulated performance parameters
    %of a reservoir whose rock and fluid properties and well parameters are
    %supplied in the input data matrix DataDeck. The elements of DataDeck are
    %imported from the accompanying Microsoft Excel spreadsheet file (2-D Reservoir Simulator
    %DataDeck Template.xlsx) which must be in the same directory with the
    %function file. The output parameter ResPerfPara gives reservoir
    %performance parameters (average reservoir pressure, average PVT
    %properties, well production rate and cummulative production) at the end of
    %simulation. However values of this performance parameters at each
    %simulation time node are exported into a Microsoft Excel spreadsheet created by the function 
    %and located in the same directory as the function file. The name of the output spreadsheet 
    %is as defined (by the user) in the input spreadsheet and imported into string array OutputFileName. 
[DataDeck,OutputFileName] = xlsread('2-D Reservoir Simulator DataDeck Template.xlsx','B1:B24');
Lx = DataDeck(1);
Ly = DataDeck(2);
h = DataDeck(3);
nx = DataDeck(4);
ny = DataDeck(5);
kx = DataDeck(6);
ky = DataDeck(7);
poro = DataDeck(8);
Swi = DataDeck(9);
Pi = DataDeck(10);
Pb = DataDeck(11);
co = DataDeck(12);
cr = DataDeck(13);
cw = DataDeck(14);
visc = DataDeck(15);
Boi = DataDeck(16);
Bob = DataDeck(17);
qt = DataDeck(18);      
WellLength = DataDeck(19);
WellHeel_XCoordinate = DataDeck(21); 
WellHeel_YCoordinate = DataDeck(22);
delta_t = DataDeck(23);
delta_x = Lx/nx; %gridblock size in x-direction
delta_y = Ly/ny; %gridblock size in y-direction
Area_x = delta_y*h; %cross-sectional area perpendicular to flow in x-direction
Area_y = delta_x*h; %cross-sectional area perpendicular to flow in y-direction
Vb = delta_x*delta_y*h; %gridblock bulk volume in cubic feet
Vb_barrel = Vb/5.615; %gridblock bulk volume in barrels
STOIP = (Lx*Ly*h*poro*(1-Swi))/(5.615*Boi); %Stock-Tank Oil initially in place in the reservoir
Ni = (delta_x*delta_y*h*poro*(1-Swi))/(5.615*Boi); %Stock-Tank Oil initially in place in each block
ce = cr+(cw*Swi)+(co*(1-Swi)); %effective compressibility of the system (rock, oil and water)
Tx = (0.001127*kx*Area_x)/(visc*Boi*delta_x); %x-direction transmissibility
Ty = (0.001127*ky*Area_y)/(visc*Boi*delta_y); %y-direction transmissibility
      %The flow model has been rearranged thus: S*P_(i,j-1) + W*P_(i-1,j) + C*P_(i,j) +
      %E*P_(i+1,j) + N*P_(i,j+1) = -G*P_(n) - qsc_(i,j)
      %S, W, C, E, N and G are coefficients of pressure terms in the model
      %qsc is the source/sink term - only applicable to wellblocks.
      %i and j are block position indices in the x- and y-direction while n is
      %time step index
G = (Vb*poro*(1-Swi)*(co+cr))/(5.615*Bob*delta_t); %Coefficient of the pressure (of the block of interest) on the RHS of the flow model
S = Ty; %Coefficient of pressure of the South neighbouring block
N = Ty; %Coefficient of pressure of the North neighbouring block
W = Tx; %Coefficient of pressure of the West neighbouring block
E = Tx; %Coefficient of pressure of the East neighbouring block
      %The coefficient matrix M is generated by generating and summing up
      %sub-matrices (five of them - one for each of the five pressure terms on the LHS of the flow model) 
      %that made up the coefficient matrix. Matrix M and its
      %sub-matrices all have dimensions nx*ny by nx*ny. Each numbered row (first index) in any
      %of the matrices represents the flow equation of the correspondingly 
      %numbered gridblock (i.e. block of interest) while each numbered column (second index) of any of the matrices  
      %represents the coefficient of the pressure of the correspondingly numbered block in 
      %the flow equation of the block of interest. Expectedly, the flow equation for a certain block of interest 
      %will only feature pressure of the block and pressures of neighbouring blocks. Consequently, the coefficients
      %will only be non-zero for the block's pressure and its neighbouring blocks' pressures and be zero for all other
      %blocks. Hence, each of the sub-matrices are initially pre-allocated
      %zeros with the zeros replaced with appropriate values only for the
      %concerned neighbouring blocks while other neighbouring (unconcerned) and
      %all non-neighbouring blocks retains the zero value. With this scheme, the
      %sum of all sub-matrices gives the coefficient matrix M. Note that blocks
      %are numbered using the natural ordering.
MS = zeros(nx*ny,nx*ny); %Sub-matrix for the coefficient of South pressure terms (i.e. P_(i,j-1) terms)                        
for I=nx+1:nx*ny %No South neighbour for Blocks 1 to nx as they are at the south boundary of the reservoir.                                   
    MS(I,I-nx)=S; %The indicated indices accurately pinpoints the Block I-nx pressure term (P_(i,j-1) term) in the flow equation for Block I
                  %Indeed, the South neighbour of Block I must neccessarily
                  %be at position I-nx in the reservoir grid model.
end                                               
MW = zeros(nx*ny,nx*ny); %Sub-matrix for the coefficient of West pressure terms (i.e. P_(i-1,j) terms)                         
for I=2:nx*ny %Block 1 is left out for it has no West neighbour!                                     
    if rem(I-1,nx)==0 %All blocks on the West Boundary of the reservoir have no West neighbour, and have to be left out. The 
                       %positions (I) of each of these blocks are 1 + multiples of nx; 
                       %hence I-1 will neccessarily be multiple on nx,
                       %consequently, dividing I-1 by nx will leave zero as remainder;
                       %the condition of the 'if' statement therefore
                       %pinpoints the blocks on the West reservoir boundary
                       %and by the reason of absence of statement between
                       %the 'if' and the 'else', these blocks are left out
                       %of the assignment commanded on the next line after
                       %the 'else' line.            
    else                                          
    MW(I,I-1)=W;       %All other blocks being assigned coefficient of West pressure term.                           
    end                %The indicated indices accurately pinpoints the Block I-1 pressure term (P_(i-1,j) term) in the flow equation for Block I
                       %Indeed, the West neighbour of Block I must neccessarily
                       %be at position I-1 in the reservoir grid model.                           
end                                               
MC = zeros(nx*ny,nx*ny); %Sub-matrix for the coefficient of Central (block of interest) pressure terms (i.e. P_(i,j) terms)
                         %Depending on the block of interest, coefficient
                         %of the Central pressure term will feature some or
                         %all of S, W, E and N and will neccesarily feature
                         %G. Note that all coefficient of the Central
                         %pressure terms (i.e. P_(i,j) terms) will
                         %neccessarily be negative.
for I=1:nx*ny
    if I == 1                                
       MC(I,I)=-(E+N+G); %Block 1 (at the bottom left corner of the reservoir grid) has only East and North neighbours 
    elseif I == nx       
       MC(I,I) = -(W+N+G); %Block nx (at the bottom right corner of the reservoir grid) has only West and North neighbours
    elseif I == nx*ny      
       MC(I,I) = -(S+W+G); %Block nx*ny (at the top right corner of the reservoir grid) has only South and West neighbours
    elseif I == (nx*(ny-1))+1  
       MC(I,I) = -(S+E+G);     %Block nx*(ny-1)+1 (at the top left corner of the reservoir grid) has only South and East neighbour.
    elseif I < nx            %This pinpoints blocks on the reservoir South boundary (except Block 1, which has been handled with an
                             %earlier 'elseif' statement and Block nx which
                             %is out of the range specified.
       MC(I,I) = -(W+E+N+G);  %These blocks lack South neighbours hence S is not included.
    elseif rem(I,nx) == 0     %This pinpoints blocks on the reservoir East boundary with position index (I) being multiples of nx so
                              % dividing I by nx gives zero as remainder
       MC(I,I) = -(S+W+N+G);  %These blocks lack East neighbours hence E is not included.
    elseif rem(I-1,nx) == 0   %This pinpoints blocks on the reservoir West boundary with position index (I) being 1 + multiples of nx so
                              % dividing I-1 by nx gives zero as remainder
       MC(I,I) = -(S+E+N+G);  %These blocks lack West neighbours hence W is not included.
    elseif I>(nx*(ny-1))+1    %This pinpoints blocks on the reservoir North boundary (except Block nx*ny which has been handled with an
                             %earlier 'elseif' statement and Block nx*(ny-1)+1 which
                             %is out of the range specified and has been handled with an
                             %earlier 'elseif' statement.
      MC(I,I) = -(S+W+E+G);  %These blocks lack North neighbours hence N is not included.
    else                     %All other blocks left (after the foregoing 'elseif' lines are actually fully interior
                             % blocks, lying on no reservoir boundary,
      MC(I,I) = -(S+W+E+N+G); % hence they have neighbours in all directions
    end                                          
end                                              
ME = zeros(nx*ny,nx*ny); %Sub-matrix for the coefficient of East pressure terms (i.e. P_(i+1,j) terms)                        
for I=1:nx*ny                                    
    if rem(I,nx)==0   %All blocks on the East Boundary of the reservoir have no East neighbour, and have to be left out. The 
                       %positions (I) of each of these blocks are multiples
                       %of nx; consequently, dividing I by nx will leave zero as remainder;
                       %the condition of the 'if' statement therefore
                       %pinpoints the blocks on the East reservoir boundary
                       %and by the reason of absence of statement between
                       %the 'if' and the 'else', these blocks are left out
                       %of the assignment commanded on the next line after
                       %the 'else' line.                            
    else                                         
    ME(I,I+1)=E;       %All other blocks being assigned coefficient of West pressure term.                          
    end                %The indicated indices accurately pinpoints the Block I+1 pressure term (P_(i+1,j) term) in the flow equation for Block I
                       %Indeed, the East neighbour of Block I must neccessarily
                       %be at position I+1 in the reservoir grid model.                          
end                                              
  MN = zeros(nx*ny,nx*ny); %Sub-matrix for the coefficient of North pressure terms (i.e. P_(i,j+1) terms)                       
for I=1:(ny-1)*nx          %This range has to be terminated at Block (ny-1)*nx thereby excluding all blocks on the reservoir
                           %North boundary which have no North neighbours.
    MN(I,I+nx)=N;          %The indicated indices accurately pinpoints the Block I+nx pressure term (P_(i,j+1) term) in the flow equation for Block I
                           %Indeed, the North neighbour of Block I must neccessarily
                           %be at position I+nx in the reservoir grid model.
end  
M=MS+MW+MC+ME+MN;         %Summing all submatrices to obtain the coefficient matrix M

      %Here the blocks hosting the horizontal well are identified and the
      %total well production rate prorated among the blocks according to
      %the lenght of well hosted by each block. The only user-defined
      %inputs into this scheme are the co-ordinate (x,y) of the location of the heel of the horizontal 
      %well in terms of its distance away (offset) from the reservoir west and south boundaries 
      %respectively as well as the length of the horizontal well (heel to toe).  
ithGridofHeel = (fix(WellHeel_XCoordinate/delta_x)) + 1; %Identifying the column of the reservoir grid where the heel is located 
                                                         %(i.e i index; note: movement along x-axis(indexed as 'i') is movement across columns)
                                                         %The identification is done by finding the number of gridblocks that needed to be
                                                         %fully covered in x-direction
                                                         %in order to move from the west boundary to the x-coordinate of 
                                                         %the heel. This number must neccessarily be the integer part (obtained using the 'fix' function)
                                                         %of x-coordinate divided by gridlength (delta_x). The +1 is to indicate that the first wellblock is 
                                                         %the block immediately after the last fully-covered block in the journey to the x-coordinate.
                                                        
jthGridofHeel = (fix(WellHeel_YCoordinate/delta_y)) + 1; %Identifying the row of the reservoir grid where the heel is located 
                                                         %(i.e j index; note: movement along y-axis(indexed as 'j') is movement across rows)
                                                         %The identification is done in a similar fashion as in ithGridofHeel
FirstWB = ((jthGridofHeel-1)*nx)+ ithGridofHeel;         %With both i and j indices of the block hosting the heel (first wellblock) known,
                                                         %the natural ordering of the first wellblock is determined using the position rule.
WL_FirstWB = delta_x-rem(WellHeel_XCoordinate,delta_x);  %In cases where the x-coordinate is not a multiple of gridlenght, the offset will occupy
                                                         %part of the first
                                                         %wellblock, hence, the segment of the welllength in first wellblock (needed for rate 
                                                         %proration) would be the gridlenght less the remainder ('rem') (after the last fully-covered 
                                                         %block) of offset shooting into first wellblock.In cases where the offset is a multiple of
                                                         %gridlenght, the 'rem' part is zero and does no harm!!!
                                                        
if rem((WellLength-WL_FirstWB),delta_x) == 0             %If the remaining part of the welllenght (after the part residing in first wellblock) is a 
   WL_LastWB = delta_x;                                  %multiple of the gridlenght, then, all remaining well blocks must be completely spanned 
   NumberofWB = 1+((WellLength-WL_FirstWB)/delta_x);     %by the welllenght so that the segment of welllenght in the last wellblock (needed 
                                                         %for rate proration) will be equal to the gridlenght. Also, the total number of 
                                                         %wellblocks would be 1 (for the first wellblock) plus the number of blocks required 
                                                         %to contain the remaining welllenght not contained by first wellblock, this number 
                                                         %is the division of remaining welllenght by gridlenght.
else
    WL_LastWB = rem((WellLength-WL_FirstWB),delta_x);    %If on the other hand the remaining part of the welllenght (after the part residing 
                                                         %in first wellblock) is not a multiple of the gridlength, then the segment of the welllength
                                                         %in the last wellblock (needed for rate proration) will be the remainder ('rem') if 
                                                         %the remaining welllenght (after welllenght in first wellblock) is divided by 
                                                         %the gridlenght. Also, between the last and first wllblocks would be a number of 
    NumberofWB = fix((WellLength-WL_FirstWB)/delta_x)+2; %wellblocks fully covered by welllenght segments; that number will be the integer part 
                                                         %when the remaining welllenght (after first wellblock) is divided by the gridlenght; 
                                                         %so that the total number of wellblocks would be that number plus 2 (1 for first wellblock 
                                                         %and 1 for last wellblock).
end

LastWB = FirstWB + NumberofWB - 1;                       %Natural ordering (position) of the last wellblock. Starting from the position of the first 
                                                         %wellblock, we simply increment the position to the tune of number of wellblocks less one 
                                                         %(less one because we need not increment for first wellblock again since we start counting from it) 
Mq = zeros(nx*ny,1);                                     %Pre-allocating a column matrix to hold the source/sink term of the flow model
    
    %Here the production rate is prorated among all wellblocks depending on
    %the welllenght segment hosted by each wellblock; the resulting
    %prorated rates are assigned to the appropriate position on the
    %source/sing matrix; other blocks not hosting welllenght segments
    %retain zeros in the matrix.
for I = FirstWB:LastWB
    if I == FirstWB
         Mq(I) = WL_FirstWB*(qt/WellLength);
    elseif I == LastWB
      Mq(I) = WL_LastWB*(qt/WellLength);
    else
        Mq(I) = delta_x*(qt/WellLength); %Except for the first wellblock and the last wellblock, the welllenght segment 
                                         %in all blocks must neccessarily be equal to the gridlenth.
    end
end
Mq;
R = ones(nx*ny,1);             %Pre-allocating a column matrix for the entire RHS of the flow model.
Plast = ones(nx*ny,1).*Pi;     %Pre-allocatting a column matrix to hold pressure of all blocks in 
                               %the nth time step as inputs (on the RHS) for solving for pressures of all blocks in the (n+1)th time step.
Pavg = Pi;                     %Creating a place holder for average reservoir pressure at the end of each time step. 
                               %Expectedly, Pi is assigned to it at initial time (t = 0),
J = 0;                         %Intializing the simulation cycle counter. Each advancement from nth time to (n+1)th time is a cycle, 
                               %starting from zero at t = 0
T = 0;                         %Initializing time
SimCycle(1,1) = 0;             %Creating and initializing (as zero) the 'Cycle Number' column of the output file. 
T_Column(1,1) = 0;             %Creating and initializing (as zero) the 'Time' column of the output file.
Np_Column(1,1) = 0;            %Creating and initializing (as zero) the 'Cummulative Production' column of the output file.
Rate_Column(1,1) = 0;          %Creating and initializing (as zero) the 'Well Rate' column of the output file.
Pressure_Column(1,1) = Pi;     %Creating and initializing (as Pi) the 'Average Reservoir Pressure' column of the output file.
OilFVF_Column(1,1) = Boi;      %Creating and initializing (as Boi) the 'Average Oil FVF' column of the output file.
BlockPressuresHead(1,1) = cellstr(char('Day 0'));
TerminatedBlockPressuresHead(1,1) = cellstr(char('Days 0'));
BlockPressures(1:nx*ny,1) = ones(nx*ny,1).*Pi;
TerminatedBlockPressures(1:nx*ny,1) = ones(nx*ny,1).*Pi;
     %Here is the solution to the model obtained, and variables are updated
     %for the next simulation cycle.
while Pavg > Pb       %The simulator is for single-phase fluid flow and so must only run while 
                      %Average reservoir pressure is yet to drop to bubble point.
    J = J+1;          %Incrementing simulation cycle counter - initialized to zero outside this loop.
    T = T+delta_t;    %Incrementing Time counter - initialized to zero outside this loop.
    for I = 1:nx*ny   
        R(I) = Mq(I)-(Plast(I)*G);  %Working out the complete RHS of the flow model, and assigning to the RHS matrix.  
    end
    R;
    Pnow = M\R;      %Solving the set of linear equations - the solution set which is pressures of 
                     %each blocks at time n+1 is stored in a column matrix Pnow;
    BlockPressuresHead(1,J+1) = cellstr(char(['Day ',num2str(T)]));
    BlockPressures(1:nx*ny,J+1) = Pnow;
    Bo = ones(nx*ny,1);  %Pre-allocating a column matrix to hold the oil FVF for each block at time n+1
    Np = zeros(nx*ny,1); %Pre-allocating a column matrix to hold the cummulative production for each block at time n+1
    Nremain = ones(nx*ny,1); %Pre-allocating a coulmn matrix to hold voulme of oil remaining in each block at time n+1 
                             %- neccessary to compute average reservoir pressure.
    NremaintimesPressure = ones(nx*ny,1); %Pre-allocating a coulmn matrix sum of which elements will be the numerator of the average reservoir pressure equation
    NremaintimesBo = ones(nx*ny,1);       %Pre-allocating a coulmn matrix sum of which elements will be the numerator of the average oil FVF equation
   for I = 1:nx*ny
      Bo(I) = Bob*(1-(co*(Pnow(I)-Pb)));   %Updating the oil FVF in each block to reflect pressure at n+1
    Np(I) = (Vb_barrel*poro*(1-0)*ce*(Pi-Pnow(I)))/Bo(I);  %Computing the cummulative oil produced from each block, at time n+1
    Nremain(I) = Ni-Np(I); %Assigning values into matrix holding stock-tank oil remaining in each block, at time n+1
    NremaintimesPressure(I) = Nremain(I)*Pnow(I);  %Assigning values into matrix sum of which elements will be the numerator of the average reservoir pressure equation
    NremaintimesBo(I) = Nremain(I)*Bo(I);          %Assigning values into matrix sum of which elements will be the numerator of the average oil FVF equation.
   end
   Npt = sum(sum(Np)); %Total reservoir cummulative production at time n+1 being sum of production from each block.
   qt_calc = Npt/T;    %Well production rate
   Pavg = (sum(sum(NremaintimesPressure)))/(STOIP-Npt); %Computing average reservoir pressure
   Boavg = (sum(sum(NremaintimesBo)))/(STOIP-Npt);   %Computing average oil FVF
  
   %Having obtained all outputs parameters, they are here assigned into the
   %appropriate row of their respective column matrices that will be
   %written to the output file. However, before they are assigned, there is
   %need to check if the average reservoir pressure is not below bubble
   %point. This check becomes neccessary considering the fact that 
   %upon receiving the output from the previous loop run, the 'while' loop
   %only stops if the average reservoir pressure is checked and found to be below
   %bubble point; before the check, the outputs are to be already written; so in
   %order to prevent writing a set of outputs that already fall below
   %bubble point, another check is made inside the 'while' loop. If the
   %latest output corresponds to pressure below bubble point, the assigning
   %is not done, i.e. even though the outputs are already generated, they are
   %discarded. If on the other hand, the outputs correponds to pressures
   %still above bubble point, the outputs are assigned to their respective
   %column matrices.
   
   %Note that in all assignments below, the Jth output is assigned to the
   %(J+1)th row of the column matrices; this is because the first row of
   %each matrix hosts the column heading.
   if Pavg > Pb
       SimCycle(J+1,1) = J;
      T_Column(J+1,1) = T;
      Np_Column(J+1,1) = Npt;
      Rate_Column(J+1,1) = qt_calc;
      Pressure_Column(J+1,1) = Pavg;
      OilFVF_Column(J+1,1) = Boavg;
      TerminatedBlockPressuresHead(1,J+1) = cellstr(char(['Day ',num2str(T)]));
      TerminatedBlockPressures(1:nx*ny,J+1) = Pnow;
   else
   end
   Plast = Pnow; %Advancing the simulation by setting the (n+1)th pressures 
                 %to nth pressures as input for the next cycle
end

    %Writing the output spreadsheet file.
OutputHead = char('Cycle Number','Time (Days)', 'Average Reservoir Pressure (psi)','Average Oil FVF', 'Cummulative production (STB)','Well Rate (STB/D)'); %Constructing the headings of the output file.
OutputFileName = num2str(char(OutputFileName)); %Putting the output file name in the format acceptable by 'xlswrite' command.
xlswrite(OutputFileName,(cellstr(OutputHead))') %Writing the output file headings
xlswrite(OutputFileName,[SimCycle, T_Column, Pressure_Column, OilFVF_Column, Np_Column, Rate_Column],1,'A2'); %Writing the output values, starting from row 2 of the spreadsheet since row 1 is occupied by the headings.
xlswrite('Block Pressures',BlockPressuresHead,1)
xlswrite('Block Pressures Terminated',TerminatedBlockPressuresHead,1)
xlswrite('Block Pressures',BlockPressures,1,'A2')
xlswrite('Block Pressures Terminated',TerminatedBlockPressures,1,'A2')

   %Some plots
figure('Name','Plot of Average Oil FVF versus Pressure', 'NumberTitle', 'off') %Creating graph figure
plot(Pressure_Column,OilFVF_Column) %Plotting on the just-created figure
title('Plot of Average Oil FVF versus Pressure')
figure('Name','Plot of Average Oil FVF versus Time', 'NumberTitle', 'off')%Creating graph figure
plot(T_Column,OilFVF_Column) %Plotting on the just-created figure
title('Plot of Average Oil FVF versus Time')
figure('Name','Plot of Reservoir Pressure versus Time', 'NumberTitle', 'off')%Creating graph figure
plot(T_Column,Pressure_Column) %Plotting on the just-created figure
title('Plot of Reservoir Pressure versus Time')
figure('Name','Plot of Cummulative Production versus Time', 'NumberTitle', 'off')%Creating graph figure
plot(T_Column,Np_Column) %Plotting on the just-created figure
title('Plot of Cummulative Production versus Time')
figure('Name','Production Rate versus Time', 'NumberTitle', 'off')%Creating graph figure
plot(T_Column,Rate_Column) %Plotting on the just-created figure
title('Plot of Rate versus Time')

    %Summary of reservoir performance parameters at the end of simulation.
Number_of_Cycle = char('Number of Cycles    ', num2str(SimCycle(J)));
T_final = char('Production Time    ',[num2str(T_Column(J)),' Days']);
Pavg_final = char('Average Reservoir Pressure    ',[num2str(Pressure_Column(J)),' psi']);
Boavg_final = char('Average Oil FVF    ',[num2str(OilFVF_Column(J)),' RB/STB']);
qt_calc_final = char('Production Rate    ',[num2str(Rate_Column(J)),' STB/D']);
Npt_final = char('Cummulative Production    ',[num2str(Np_Column(J)),' STB']);
ResPerfPara = [Number_of_Cycle T_final Pavg_final Boavg_final qt_calc_final Npt_final]

end