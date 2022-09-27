import * as React from 'react';
import styles from '../safetyHub/components/SafetyHub.module.scss';
//import styles from '../../SafetyHub.module.scss';

import { Spinner } from 'office-ui-fabric-react/lib/Spinner';
import { IStackTokens, Stack } from 'office-ui-fabric-react/lib/Stack';
export const Loader: React.FunctionComponent = () => {
    const stackTokens: IStackTokens = {
        childrenGap: 20,

    };

    return (
        <div className={styles.safetyHub} >
            <div>
                <Spinner label="loading...." />
            </div>
        </div>
    );
};