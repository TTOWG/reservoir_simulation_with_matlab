function NP = TTOWGbal2D(A,BlockPressure)

A = xlsread('TTOWGbal2D - Model Parameters.xlsx','B1:B15');
Lx = A(1,1);%Length of the reservior, in x-direction(ft)
Ly = A(2,1);%Lenght of the reservior, in y-direction(ft)
h = A(3,1);%Ver+tical thickness of the reservior(ft)
nx = A(4,1);%Number of blocks in x-direction
ny = A(5,1);%Number of blocks in y-direction
poro = A(6,1);%Reservior porosity (fraction)
Swi = A(7,1);%Initial Water Saturation(fraction)
Pi = A(8,1);%Initial Reservoir Pressure (psi)
Pb = A(9,1);%Bubble point Pressure (psi)
Co = A(10,1);%Oil Compressibilty(/psi)
cf = A(11,1);%Formation Compressibility(/psi)
cw = A (12,1);%Water Compressibility (/psi)
Boi = A(13,1);%Oil Formation Volume Factor at initial Pressure (RB/STB)
Bob = A(14,1);%Oil Formation Volume Factor at Bubble-point Pressure (RB/STB)
Td = A(15,1); %Time steps (days)

Ce = (((1-Swi)*Co)+cf+(Swi*cw))/(1-Swi);%Ce is the effective compressibility
dx = Lx/nx;
dy = Ly/ny;
N =(dx*dy*(1-Swi)*poro*h)/(5.615*Boi); % N is the amount of oil initially in place in each grid block
STOIIP = (Lx*Ly*(1-Swi)*poro*h)/(5.615*Boi);%Finds the Stock tank Oil initially in place for the reservior

BlockPressure = xlsread('TTOWGbal2D - Block Pressures.xls');%Importing the table of pressure from the excel file
allcolumn = size(BlockPressure,2);
nt = allcolumn-1;
TB = zeros(nx,ny);%A zero matrix with the same number of row and columns as specified by nx and ny
T=0;
PPrint(1,1)=Pi;
OilRemain(1,1)=STOIIP;
Num(1,1) = 0;
CNPPrint(1,1)= 0;
Tdays(1,1)= 0;
Rates(1,1) = 0;

for J = 1:nt %Creates a count of the values from 1 to the calculated value of nt
    P =  BlockPressure(2:((ny*nx)+1),J+1);
    for i = 1:ny
        for j = 1:nx
            TB(i,j)= P((i-1)*nx+j,1);
        end
    end
PPb= TB-Pb;%PPb is the difference between the matrix of the pressure values and the bubble point pressure
Bo = Bob*(1-(Co.*PPb));% This calculates oil formation volume factor for each block.
PiPn = Pi-TB;%This is the pressure difference (Pi-P)
Val= N*Boi*Ce;
Np = (Val.*PiPn)./Bo;%This calculates the cumulative amount of oil produced from each block, at that time
CNP = sum(Np(:));%This is the cumulative oil produced from the entire reservoir
OilRem = N-Np;%This calculating for the oil remaining in each block
Numerator = OilRem.*TB;
TNT = sum(Numerator(:));%This is the sum of the matrix of pressure values multiplied by the oil remaining
Denum = STOIIP-CNP;%Finds amount of oil remaining in the reservior
Pav = TNT/Denum;%Calculates the Average Pressure of the Reservior
T = T+Td;
Flowrate = CNP/T;

if Pav > Pb %This checks if Average reservoir Pressure is still above bubble point; otherwise, simulation resluts is not recorded.

Num(J+1,1) = J;   
PPrint(J+1,1)= Pav;
CNPPrint(J+1,1)= CNP;
Tdays(J+1,1)= T;
Rates(J+1,1) = Flowrate;
OilRemain(J+1,1)=Denum;
end
end

xlswrite('Performance Parameters 2D',{'Cycles','Time (Days)','Average Pressure (psi)','Flow Rate (STB/D)','Cummulative Oil Produced (STB)','Oil Remaining (STB)'},1,'A1');
xlswrite('Performance Parameters 2D',[Num,Tdays,PPrint,Rates,CNPPrint,OilRemain],1,'A2');

end

