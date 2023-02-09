function NP = TTOWGbal3D(A,BlockPressure)

A = xlsread('Model Parameters.xlsx');
nLayers = size(A,2);%Determining the number of layers, being the unmber of columns in the input data deck
for L = 1:nLayers
   Lx(1,L) = A(1,L);%Length of the reservior in x-direction (ft), for each layer
   Ly(1,L) = A(2,L);%Lenght of the reservior in y-direction (ft), for each layer
   h(1,L) = A(3,L);%Vertical thickness (ft), for each layer of the reservior
   nx(1,L) = A(4,L);%Number of blocks in x-direction, for each layer
   ny(1,L) = A(5,L);%Number of blocks in y-direction, for each layer
   poro(1,L) = A(6,L);%Reservior porosity (fraction), for each layer
   Swi(1,L) = A(7,L);%Initial Water Saturation(fraction), for each layer
   Pi(1,L) = A(8,L);%Initial Reservoir Pressure (psi), for each layer
   Pb(1,L) = A(9,L);%Bubble point Pressure (psi), for each layer
   Co(1,L) = A(10,L);%Oil Compressibilty(/psi), for each layer
   cf(1,L) = A(11,L);%Formation Compressibility(/psi), for each layer
   cw(1,L) = A(12,L);%Water Compressibility (/psi), for each layer
   Boi(1,L) = A(13,L);%Oil Formation Volume Factor at initial Pressure (RB/STB), for each layer
   Bob(1,L) = A(14,L);%Oil Formation Volume Factor at Bubble-point Pressure (RB/STB), for each layer
end
Td = A(15,1); %Time steps (days)
Ce = (((1-Swi).*Co)+cf+(Swi.*cw))./(1-Swi);%Ce is the effective compressibility
dimxdimy = nx.*ny; %Need the LHS in allocating pressure from the pressure vector to the 3-D grid configuration.
dx = Lx./nx;
dy = Ly./ny;
N =(dx.*dy.*(1-Swi).*poro.*h)./(5.615*Boi); % N is the amount of oil initially in place in each grid block
LayerSTOIIP = (Lx.*Ly.*(1-Swi).*poro.*h)./(5.615*Boi);%Finds the Stock tank Oil initially in place in each layer
STOIIP = sum(LayerSTOIIP);%Finds the Stock tank Oil initially in place in the entire reservior

BlockPressure = xlsread('Block Pressures Matrix.xls');%Importing the table of pressure from the excel file
nCycles = size(BlockPressure,2)-1;
T=0;
PPrint(1,1)= sum(N.*nx.*ny.*Pi)/STOIIP;
OilRemain(1,1)=STOIIP;
CycleNum(1,1) = 0;
CNPPrint(1,1)= 0;
Tdays(1,1)= 0;
Rates(1,1) = 0;

for J = 1:nCycles %Creates a count of the values from 1 to the calculated value of nCycles
    P =  BlockPressure(2:((sum(ny.*nx))+1),J+1);
    for L = 1:nLayers
        for i = 1:ny
            for j = 1:nx
                TB(i,j)= P(((sum(dimxdimy(1:(L-1))))+((i-1)*nx(L))+j),1);%3D pressure distribution for each layer.
            end
        end
        PPb= TB-Pb(L);%PPb is the difference between the matrix of the pressure values and the bubble point pressure,
        Bo = Bob(L)*(1-(Co(L).*PPb));% This calculates oil formation volume factor for each block.
        PiPn = Pi(L)-TB;%This is the pressure difference (Pi-P)
        Val= N(L)*Boi(L)*Ce(L);
        Np = (Val.*PiPn)./Bo;%This calculates the cumulative amount of oil produced from each block, at that time
        LayerCNP(L) = sum(Np(:));%This is the cumulative oil produced from a layer
        BlockOilRem = N(L)-Np;%This calculates the oil remaining in each block
        Numerator = BlockOilRem.*TB;
        LayerTNT(L) = sum(Numerator(:));%This is the sum of the matrix of pressure values multiplied by the oil remaining
        LayerOilRem(L) = LayerSTOIIP(L)-LayerCNP(L);%Finds amount of oil remaining in each layer
        LayerPav(L) = LayerTNT(L)/LayerOilRem(L);%Calculates the Average Pressure of each layer
    end
    T = T+Td;
    CNP = sum(LayerCNP);
    OilRem = sum(LayerOilRem);
    Flowrate = CNP/T;
    Pav = sum(LayerPav.*LayerOilRem)/sum(LayerOilRem);



    if Pav > Pb %This checks if Average reservoir Pressure is still above bubble point; otherwise, simulation resluts is not recorded.

      CycleNum(J+1,1) = J;   
      PPrint(J+1,1)= Pav;
      CNPPrint(J+1,1)= CNP;
      Tdays(J+1,1)= T;
      Rates(J+1,1) = Flowrate;
      OilRemain(J+1,1)=OilRem;
   end
end

xlswrite('Performance Parameters',{'Cycles','Time (Days)','Average Pressure (psi)','Flow Rate (STB/D)','Cummulative Oil Produced (STB)','Oil Remaining (STB)'},1,'A1');
xlswrite('Performance Parameters',[CycleNum,Tdays,PPrint,Rates,CNPPrint,OilRemain],1,'A2');

end

