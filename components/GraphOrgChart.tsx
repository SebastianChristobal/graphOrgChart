import * as React from 'react';
import styles from './GraphOrgChart.module.scss';
import { IGraphOrgChartProps } from './IGraphOrgChartProps';
import { OrgChart } from './OrgChart';
import { OrgChartContext } from '../context/OrgChartContext';

const GraphOrgChart:React.FC<IGraphOrgChartProps> = ({context, defaultUser,showActionsBar,showAllManagers,startFromUser,title, extensionAttributeValue, filterByCompanyName}) => {
  const contextValue = React.useMemo(() => {
    return { context, extensionAttributeValue };
  }, [context,extensionAttributeValue]);
  // const me = await graphService.GetUserByEmail("gustav.linden@lindendevelop.onmicrosoft.com");
    return (
      <section className={`${styles.graphOrgChart}`}>
        <OrgChartContext.Provider value={contextValue}>

        <OrgChart context={context} defaultUser={defaultUser} showActionsBar={showActionsBar} filterByCompanyName={filterByCompanyName} showAllManagers={showAllManagers} startFromUser={startFromUser} title={title}/>
        </OrgChartContext.Provider>
      </section>
    );
  
}
export default GraphOrgChart
